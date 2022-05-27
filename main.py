import docx
import re
def GetParaData(output_doc_name, paragraph):
    output_para = output_doc_name.add_paragraph()
    for run in paragraph.runs:
        output_run = output_para.add_run(run.text)
        # Run's bold data
        output_run.bold = run.bold
        # Run's italic data
        output_run.italic = run.italic
        # Run's underline data
        output_run.underline = run.underline
        # Run's color data
        output_run.font.color.rgb = run.font.color.rgb
        # Run's font data
        output_run.style.name = run.style.name
        # Paragraph's alignment data
        output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
        output_para.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
        output_para.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
        output_para.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
class document:
    def Serch(self,fileName):
        self.fileName=fileName
        doc = docx.Document(fileName)

        completedText = []
        BaseGost=[['ГОСТ 8','2000'],['ГОСТ 1','2002'],['ГОСТ 2','2001'],['ГОСТ 9','2003'],['ГОСТ 0','2000']]
        for paragraph in doc.paragraphs:
            completedText.append(paragraph.text)
        hub=[]
        res=[]
        result=[]
        lengthparag = len(completedText)
        for i in range(lengthparag):
                result4= re.findall(r"ГОСТ \d{1}-\d{4}", completedText[i])
                if result4!=[]:
                    hub.append(result4)

        res.append(sum(hub, []))
        print(res)
        b=[]
        print(b)
        for i in range(len(res[0])):
            g = res[0][i].split('-')
            b.append(g)
        #a = np.array(g)
        #b = a.reshape(-1, 2)
        print(b)
        #print(b)
        for i in range(len(b)):
            for j in range(len(BaseGost)):
                if b[i][0] == BaseGost[j][0]:
                    print('Same gost')
                    if b[i][1] == BaseGost[j][1]:
                        print('Same god')
                    else:
                        result.append(b[i][0])
                        print(result)
                        print('Not same god')
        for i in range(lengthparag):
                for j in range(len(result)):
                        if completedText[i].find(result[j]) != 1:
                         char = completedText[i]
                         char = char.replace(result[j], result[j]+'-ЭТОТ ГОСТ УСТАРЕЛ')
                         completedText[i] = char
                        else:
                         print('Строка ненайдена')
        return [completedText,lengthparag,doc.paragraphs]

    def Save(Paragraph, ct):
        n = len(Paragraph)
        p = docx.Document()
        for i in range(n):
            GetParaData(p, Paragraph[i])
            p.paragraphs[i].text = ct[i]
        p.save('demo2.docx')
O=document()
BigO=O.Serch('demo.docx')

document.Save(BigO[2],BigO[0])

