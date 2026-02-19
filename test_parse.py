import sys
sys.path.insert(0, r'C:\Users\Kingry\Documents\Essay Processor')
from essay_processor import EssayProcessor

# Test PDF parsing
ep = EssayProcessor.__new__(EssayProcessor)
ep.pdf_files = {}
ep.docx_files = {}
ep.output_files = {}

pdf_path = r'C:\Users\Kingry\Documents\Essay Processor\Taylor_  Light pollution_review.pdf'
data = ep.parse_pdf_feedback(pdf_path)

print('=== OVERALL GRADE ===')
print(f'Grade: {data["overall_grade"]}')
print(f'Overview: {data["overall_overview"][:200]}...')

print(f'\nNumber of sections: {len(data["sections"])}')
for sec in data['sections']:
    print(f'  - {sec["name"]}: {sec["grade"]} ({len(sec["quotes"])} quotes)')
