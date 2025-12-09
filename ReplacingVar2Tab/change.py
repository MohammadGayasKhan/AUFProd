import json
from docxtpl import DocxTemplate

def process_file(input_filename, output_filename=None):
    with open('details.json', 'r') as f:
        context = json.load(f)
    
    doc = DocxTemplate(input_filename) 
    
    doc.render(context)
    
    if output_filename is None:
        output_filename = input_filename.replace('.docx', '_output.docx')
    
    doc.save(output_filename)
    
    print(f"Processed file saved to: {output_filename}")
    return context

if __name__ == "__main__":
    context = process_file('Template.docx', 'output.docx')
    
    print("\n" + "="*50)
    print("Template variables used:", list(context.keys()))
    print("="*50)
