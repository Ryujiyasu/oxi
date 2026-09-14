import unittest
from word_symbols import restore_symbols

class SymbolTests(unittest.TestCase):
    def xml(self, body):
        return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' + body + '</w:p></w:body></w:document>'
    def test_symbol_and_literal_parentheses(self):
        xml = self.xml('<w:r><w:sym w:font="Symbol" w:char="F07F"/></w:r><w:r><w:t> Referred (Detox)</w:t></w:r>')
        self.assertEqual(restore_symbols('( Referred (Detox)\r\x07', xml), '\uf07f Referred (Detox)\r\x07')
    def test_plain_parenthesis_is_unchanged(self):
        self.assertEqual(restore_symbols('(ordinary)\r', self.xml('<w:r><w:t>(ordinary)</w:t></w:r>')), '(ordinary)\r')
    def test_mismatched_paragraph_is_unchanged(self):
        self.assertEqual(restore_symbols('( different', self.xml('<w:r><w:sym w:char="F084"/><w:t> other</w:t></w:r>')), '( different')
    def test_invalid_metadata_is_unchanged(self):
        for xml in ['invalid', self.xml('<w:r><w:sym w:char="bad-value"/></w:r>')]:
            self.assertEqual(restore_symbols('(', xml), '(')

if __name__ == '__main__': unittest.main()
