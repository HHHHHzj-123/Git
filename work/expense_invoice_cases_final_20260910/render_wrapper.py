import os,runpy,sys
os.environ['PATH']=r'C:\Program Files\LibreOffice\program;C:\Users\HZJ\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin;'+os.environ.get('PATH','')
sys.argv=[r'C:\Users\HZJ\.codex\plugins\cache\openai-primary-runtime\documents\26.905.11957\skills\documents\render_docx.py',r'C:\Users\HZJ\Desktop\Git\work\expense_invoice_cases_final_20260910\deliverables\企业发票、报销与异常支出实战手册.docx','--output_dir',r'C:\Users\HZJ\Desktop\Git\work\expense_invoice_cases_final_20260910\qa\rendered','--emit_pdf']
runpy.run_path(sys.argv[0],run_name='__main__')
