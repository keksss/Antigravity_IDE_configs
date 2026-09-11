
import sys
import win32com.client

files = ["C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\01_headings_and_hierarchy\\headings_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\02_inline_formatting\\inline_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\03_hyperlinks_and_anchors\\links_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\04_lists_and_hierarchies\\lists_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\05_gfm_tables_and_alignment\\tables_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\06_blockquotes_and_callouts\\callouts_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\07_code_blocks_and_syntax\\code_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\08_images_and_aspect_ratio\\images_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\09_template_inheritance\\template_output.docx", "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\10_chained_workflow\\contract_stage1.docx"]
pdf_target = "C:\\Users\\kekss\\Project\\Antigravity_IDE_configs\\tests\\docx-tools\\markdown-to-docx\\scenarios\\11_word_com_automation\\sample_validation.pdf"

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    for idx, f in enumerate(files):
        doc = word.Documents.Open(f, ReadOnly=True, ConfirmConversions=False)
        print(f"Opened {idx+1}/{len(files)}: {f}")
        if idx == 0:
            doc.SaveAs(pdf_target, FileFormat=17) # 17 = wdFormatPDF
            print(f"PDF exported successfully to {pdf_target}")
        doc.Close(SaveChanges=False)
        
    word.Quit()
    print("ALL_DOCS_PASSED_WORD_COM")
except Exception as e:
    print("COM_ERROR:", e, file=sys.stderr)
    try:
        word.Quit()
    except Exception:
        pass
    sys.exit(1)
