"""Recover OOXML symbol codes from Word's literal-parenthesis placeholders."""
import xml.etree.ElementTree as ET

_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def restore_symbols(raw_text: str, word_xml: str) -> str:
    """Change only symbols whose complete paragraph matches Word's raw text."""
    try:
        root = ET.fromstring(word_xml)
        for paragraph in root.findall(".//w:body//w:p", _NS):
            if not paragraph.findall(".//w:sym", _NS):
                continue
            plain, restored = [], []
            for node in paragraph.iter():
                tag = node.tag.rsplit("}", 1)[-1]
                if tag == "t":
                    plain.append(node.text or "")
                    restored.append(node.text or "")
                elif tag == "sym":
                    value = chr(int(node.attrib["{" + _NS["w"] + "}char"], 16))
                    plain.append("(")
                    restored.append(value)
                elif tag in ("tab", "br"):
                    value = "\t" if tag == "tab" else "\x0b"
                    plain.append(value)
                    restored.append(value)
            expected = "".join(plain)
            if expected == raw_text.rstrip("\r\x07"):
                return "".join(restored) + raw_text[len(expected):]
    except (ET.ParseError, KeyError, ValueError, OverflowError):
        pass
    return raw_text
