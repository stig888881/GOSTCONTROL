from docx import Document
import re
import numpy as np
class document:
    @staticmethod
    def GetParaData(output_doc_name, paragraph):
        output_para = output_doc_name.add_paragraph()
        for run in paragraph.runs:
            output_run = output_para.add_run(run.text)
            output_run.bold = run.bold
            output_run.italic = run.italic
            output_run.underline = run.underline
            output_run.font.color.rgb = run.font.color.rgb
            output_run.style = run.style
            output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
            output_para.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
            output_para.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
            output_para.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
            output_para.content_type = paragraph.part.content_type
    @staticmethod
    def BaseGost():
        import sqlite3
        conn = sqlite3.connect("sqlite.db")
        gost2 = []
        conn.row_factory = lambda cursor, row: row[0]
        c = conn.cursor()
        gost2 = c.execute('SELECT * FROM tablegost2').fetchall()
        m = []
        for i in range(len(gost2)):
            g = gost2[i].split('-')
            m.append(g[0])
            m.append(g[1])
        a = np.array(m)
        b = a.reshape(-1, 2)
        resgost2 = np.array(b).tolist()
        return resgost2

    def Serch(self,fileName):
        self.fileName=fileName
        doc = Document(fileName)

        completedText = []
        BaseGost=self.BaseGost()
        #BaseGost=[['ГОСТ Р 8','2000'],['ГОСТ Р 1','2002'],['ГОСТ Р 2','2001'],['ГОСТ Р 9','2003'],['ГОСТ Р 0','2000']]
        for paragraph in doc.paragraphs:
            completedText.append(paragraph.text)
        hub=[]
        res=[]
        result=[]
        lengthparag = len(completedText)
        for i in range(lengthparag):
                result4= re.findall(r"ГОСТ Р \d{5}-\d{2}", completedText[i])
                if result4!=[]:
                    hub.append(result4)

        res.append(sum(hub, []))
        print(res)
        b=[]
        print(b)
        for i in range(len(res[0])):
            g = res[0][i].split('-')
            b.append(g)
        print(b)
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
                         doc.paragraphs[i].text=char
                        else:
                         print('Строка ненайдена')
        return doc.paragraphs

    def Save(self,Paragraph):
        n = len(Paragraph)
        p = Document()
        for i in range(n):
            self.GetParaData(p, Paragraph[i])
        p.save('demo2.docx')
O=document()
BigO=O.Serch('demo.docx')
O.Save(BigO)

