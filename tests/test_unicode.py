"""Unicode / encoding guarantees for message conversion and folder names.

Every converted message must be pure ASCII on the wire (so IMAP APPEND never
needs 8BITMIME), and parsing it back must yield exactly the original Unicode
subject, attachment filenames and body.
"""

import base64
import unittest
from email import message_from_string
from email.header import decode_header, make_header

from tests import support

mig = support.mig


def utf7_decode(text):
    """Minimal modified UTF-7 decoder for round-trip tests."""
    out = []
    index = 0
    while index < len(text):
        if text[index] == "&":
            if index + 1 < len(text) and text[index + 1] == "-":
                out.append("&")
                index += 2
                continue
            end = text.index("-", index)
            b64 = text[index + 1:end].replace(",", "/")
            b64 += "=" * (-len(b64) % 4)
            out.append(base64.b64decode(b64).decode("utf-16-be"))
            index = end + 1
        else:
            out.append(text[index])
            index += 1
    return "".join(out)


def build_message(subject, body, attachments=()):
    message = {
        "id": "m1",
        "subject": subject,
        "from": {"emailAddress": {"address": "a@b.com"}},
        "toRecipients": [{"emailAddress": {"address": "c@d.com"}}],
        "ccRecipients": [],
        "bccRecipients": [],
        "receivedDateTime": "2024-01-02T03:04:05Z",
        "body": {"contentType": "html", "content": body},
    }
    if attachments:
        message["attachments"] = [
            {"name": name, "contentBytes": base64.b64encode(content).decode()}
            for name, content in attachments
        ]
    return message


def decoded_subject(raw):
    return str(make_header(decode_header(raw.get("Subject", ""))))


class ConversionUnicodeTests(unittest.TestCase):
    CASES = [
        # (subject, body, [(filename, content)])
        ("RE: Opravená zálohová faktura – Zámek Mělník",
         "<p>Ahoj světe 🤝 中文 日本語 čeština</p>",
         [("příloha č. 1.pdf", "PDF-obsah".encode("utf-8"))]),
        ("日本語の件名 🤝", "<p>テスト本文</p>",
         [("添付ファイル.docx", "doc".encode())]),
        ("Emoji subject 🤝🚀 and CJK 中文", "plain body with 🤝",
         [("emoji 🤝.png", b"\x89PNG")]),
        ("ASCII only", "<p>hello</p>", [("file.txt", b"hi")]),
        ("Quotes \" and / slash in subject",
         "<p>x</p>", [('weird"name.pdf', b"x"), ("a/b\\c.txt", b"y")]),
        ("Dlouhá česká věta s diakritikou ěščřžýáíéůú ďťň ó é á í ý ž",
         "<p>žluťoučký kůň úpěl ďábelské ódy</p>", []),
        ("None-ish", "<p></p>", []),
    ]

    def setUp(self):
        support.install_fakes()
        self.cfg = support.make_cfg(migrate_attachments=True)

    def test_subjects_filenames_bodies_round_trip(self):
        for subject, body, attachments in self.CASES:
            with self.subTest(subject=subject):
                message = build_message(subject, body, attachments)
                rfc822 = mig.convert_to_rfc822(self.cfg, message)
                self.assertIsInstance(rfc822, str)
                self.assertTrue(rfc822.isascii(),
                                f"non-ASCII leaked into output for {subject!r}")

                parsed = message_from_string(rfc822)
                self.assertEqual(decoded_subject(parsed).strip(), subject.strip())

                parsed_files = sorted(
                    part.get_filename() for part in parsed.walk()
                    if part.get_filename())
                # Filenames are sanitized: path separators and quotes become "_".
                expected_files = sorted(
                    mig.attachment_filename(name) for name, _ in attachments)
                self.assertEqual(parsed_files, expected_files)

                html_part = next(part for part in parsed.walk()
                                 if part.get_content_type() == "text/html")
                self.assertEqual(html_part.get_payload(decode=True).decode("utf-8"),
                                 mig.extract_email_body(message, "html"))

    def test_plain_format_also_round_trips(self):
        cfg = support.make_cfg(migrate_attachments=False, email_format="plain")
        message = build_message("Příliš žluťoučký kůň", "<p>Ďábelské ódy 🤝</p>")
        rfc822 = mig.convert_to_rfc822(cfg, message)
        self.assertTrue(rfc822.isascii())
        parsed = message_from_string(rfc822)
        self.assertEqual(decoded_subject(parsed), "Příliš žluťoučký kůň")
        plain = next(part for part in parsed.walk()
                     if part.get_content_type() == "text/plain")
        self.assertEqual(plain.get_payload(decode=True).decode("utf-8"),
                         "Ďábelské ódy 🤝")

    def test_nested_eml_keeps_unicode(self):
        nested = build_message("Vnořená zpráva – příloha 🤝", "<p>vnitřní tělo</p>",
                               [("vnitřní soubor.pdf", "vnitř".encode("utf-8"))])
        outer = build_message("Vnější", "<p>vnější</p>")
        outer["attachments"] = [{
            "@odata.type": "#microsoft.graph.itemAttachment",
            "name": "RE: vnořená.eml",
            "item": nested,
        }]
        rfc822 = mig.convert_to_rfc822(self.cfg, outer)
        self.assertTrue(rfc822.isascii())

        parsed = message_from_string(rfc822)
        eml_part = next(part for part in parsed.walk()
                        if part.get_content_type() == "message/rfc822")
        self.assertEqual(eml_part.get_filename(), "RE: vnořená.eml")
        inner = eml_part.get_payload(0)
        self.assertEqual(decoded_subject(inner), "Vnořená zpráva – příloha 🤝")
        inner_files = [part.get_filename() for part in inner.walk()
                       if part.get_filename()]
        self.assertEqual(inner_files, ["vnitřní soubor.pdf"])

    def test_long_unicode_subject_folds_and_decodes(self):
        subject = "Velmi dlouhá česká věta " * 10 + "ď"
        message = build_message(subject, "<p>x</p>")
        rfc822 = mig.convert_to_rfc822(self.cfg, message)
        self.assertTrue(rfc822.isascii())
        self.assertEqual(decoded_subject(message_from_string(rfc822)), subject)


class HtmlCharsetTests(unittest.TestCase):
    def test_meta_charset_rewritten(self):
        body = '<html><head><meta charset="windows-1252"></head><body>x</body></html>'
        self.assertIn('charset="utf-8"', mig.normalize_html_charset(body))

    def test_http_equiv_content_type_rewritten(self):
        body = ('<meta http-equiv="Content-Type" '
                'content="text/html; charset=iso-8859-2">')
        self.assertIn("charset=utf-8", mig.normalize_html_charset(body))

    def test_xml_declaration_rewritten(self):
        body = '<?xml version="1.0" encoding="windows-1252"?><p>x</p>'
        self.assertIn('encoding="utf-8"', mig.normalize_html_charset(body))

    def test_utf8_and_absent_untouched(self):
        self.assertEqual(mig.normalize_html_charset('<meta charset="utf-8">'),
                         '<meta charset="utf-8">')
        self.assertEqual(mig.normalize_html_charset("<p>plain</p>"), "<p>plain</p>")

    def test_extract_applies_rewrite_for_html_only(self):
        msg = {"body": {"contentType": "html",
                        "content": '<meta charset="windows-1252"><p>ž</p>'}}
        self.assertIn('charset="utf-8"', mig.extract_email_body(msg, "html"))
        msg_text = {"body": {"contentType": "text",
                             "content": '<meta charset="windows-1252">ž'}}
        self.assertIn("windows-1252", mig.extract_email_body(msg_text, "html"))


class FolderNameUnicodeTests(unittest.TestCase):
    NAMES = [
        "Odeslaná pošta",
        "日本語フォルダ",
        "Emoji 🤝 folder",
        "Složka & ampersand",
        "Wysłane",
        "~peter/mail/台北/日本語",
        "INBOX",
    ]

    def test_modified_utf7_round_trip(self):
        for name in self.NAMES:
            with self.subTest(name=name):
                encoded = mig.imap_utf7_encode(name)
                self.assertTrue(encoded.isascii())
                self.assertEqual(utf7_decode(encoded), name)

    def test_rfc_vector(self):
        self.assertEqual(mig.imap_utf7_encode("~peter/mail/台北/日本語"),
                         "~peter/mail/&U,BTFw-/&ZeVnLIqe-")

    def test_target_path_keeps_unicode_and_neutralises_delimiter(self):
        cfg = support.make_cfg(target_parent_folder="Migration")
        path = mig.build_target_path(cfg, "Odeslaná pošta", "/")
        self.assertEqual(path, "Migration/Odeslaná pošta")
        # A display name containing the delimiter arrives as separate path
        # components (Graph paths are "/"-joined), and quotes are stripped.
        path = mig.build_target_path(cfg, 'A/B "quoted"', "/")
        self.assertEqual(path, "Migration/A/B quoted")
        # ...but a literal delimiter character must not survive on dot-servers.
        path = mig.build_target_path(support.make_cfg(), 'A.B "q"', ".")
        self.assertEqual(path, "A_B q")
        # Dot-delimited servers
        path = mig.build_target_path(support.make_cfg(), "Odeslaná/pošta", ".")
        self.assertEqual(path, "Odeslaná.pošta")

    def test_attachment_filename_sanitisation_keeps_unicode(self):
        self.assertEqual(mig.attachment_filename("příloha č.1.pdf"), "příloha č.1.pdf")
        self.assertEqual(mig.attachment_filename('a/b\\c"d\x00e.txt'), "a_b_c_d_e.txt")
        self.assertEqual(mig.attachment_filename("", forced_suffix=".eml"),
                         "attachment.eml")
        self.assertEqual(mig.attachment_filename("zprava", forced_suffix=".eml"),
                         "zprava.eml")
        self.assertEqual(mig.attachment_filename("zprava.eml", forced_suffix=".eml"),
                         "zprava.eml")


if __name__ == "__main__":
    unittest.main()
