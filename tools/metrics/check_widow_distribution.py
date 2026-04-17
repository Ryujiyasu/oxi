"""Check widow_control distribution across regressed and improved docs."""
import win32com.client, os, json

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

regressed_docs = [
    "04b88e7e0b25_index-19",
    "459f05f1e877_kyodokenkyuyoushiki01",
    "d77a58485f16_20240705_resources_data_outline_08",
    "ed025cbecffb_index-23",
    "b5f706e9f6ad_kyodokenkyuyoushiki_bessi",
    "gen2_077_API_Documentation",
    "gen2_014_出張報告書",
]

improved_docs = [
    "b837808d0555_20240705_resources_data_guideline_02",
    "c7b923e5c616_20240705_resources_data_outline_06",
]

DOCX_DIR = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx"

def check(doc_name):
    path = os.path.join(DOCX_DIR, doc_name + ".docx")
    if not os.path.exists(path):
        # Try find by prefix
        import glob
        matches = glob.glob(os.path.join(DOCX_DIR, doc_name[:20] + "*.docx"))
        if matches:
            path = matches[0]
        else:
            print(f"  NOT_FOUND: {doc_name}")
            return
    wdoc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    try:
        n = wdoc.Paragraphs.Count
        wc_false = 0
        wc_true = 0
        for i in range(1, min(n, 50) + 1):
            p = wdoc.Paragraphs(i)
            if p.Format.WidowControl:
                wc_true += 1
            else:
                wc_false += 1
        total = wc_false + wc_true
        print(f"  {doc_name}: widow_false={wc_false}/{total}, true={wc_true}/{total}")
    finally:
        wdoc.Close(False)

print("REGRESSED:")
for d in regressed_docs:
    check(d)
print("\nIMPROVED:")
for d in improved_docs:
    check(d)

word.Quit()
