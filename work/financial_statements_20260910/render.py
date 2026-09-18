import os,runpy,sys
os.environ['PATH']=r'C:\Program Files\LibreOffice\program;C:\Users\HZJ\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin;'+os.environ.get('PATH','')
sys.argv=[r'C:\Users\HZJ\.codex\plugins\cache\openai-primary-runtime\documents\26.905.11957\skills\documents\render_docx.py',r'C:\Users\HZJ\Desktop\Git\work\financial_statements_20260910\deliverables\企业财务报表编制、调整、重分类与校验实操手册.docx','--output_dir',r'C:\Users\HZJ\Desktop\Git\work\financial_statements_20260910\qa\rendered','--emit_pdf']
runpy.run_path(sys.argv[0],run_name='__main__')
