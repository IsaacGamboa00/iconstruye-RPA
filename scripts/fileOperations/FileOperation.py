import glob
import os
import smtplib
import win32com.client as win32
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText




class Operation:

    def __init__(self, downloadPath):
        self.downloadPath = downloadPath
        self.ruta_adjunto = "./files/"

    def getFileName(self):
        files_path = os.path.join(self.downloadPath, '*')
        files = sorted(glob.iglob(files_path), key=os.path.getctime, reverse=True)
        return (os.path.basename(files[0]))

    def renameFile(self, fileName):
        dirContents = os.listdir(self.downloadPath)
        if len(dirContents) == 0:
            print('no hay archivos para renombrar')
        else:
            os.rename(self.downloadPath + "" + self.getFileName(), self.downloadPath + "" + fileName)

    def removeFile(self):
        dirContents = os.listdir(self.downloadPath)
        if len(dirContents) == 0:
            print('Folder is Empty')
        else:
            print("se elimino un archivo llamado: " + self.downloadPath + "" + self.getFileName())
            os.remove(self.downloadPath + "" + self.getFileName())


    def convertFormatToXls(self,name):
        
        dirContents = os.listdir(self.downloadPath)
        if len(dirContents) == 0:
            print('no hay archivos para convertir')
        else:
            fname = self.downloadPath + "" + self.getFileName()
           
   
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            print("se convirtio a: "+name+".xlsx")
            wb.SaveAs(self.downloadPath+name+".xlsx", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()                             # FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            
    def convertFormatToCsv(self,name):
        
        data_xls = pd.read_excel(self.downloadPath + "" + self.getFileName(), index_col=0)
        data_xls.to_csv(self.downloadPath +name+'.csv', encoding='utf-8')   
        

    def sendEmail(self, asunto, cuerpo):
        dirContents = os.listdir(self.downloadPath)
        if len(dirContents) == 0:
            print('no hay archivos para enviar')
        else:
            remitente = 'iconstruyemkingenieria@gmail.com'
            destinatarios = ['suti.pruebas.bot@gmail.com']
            password = 'Suti0220*'
            mensaje = MIMEMultipart()
            mensaje['From'] = remitente
            mensaje['To'] = ", ".join(destinatarios)
            mensaje['Subject'] = asunto

            mensaje.attach(MIMEText(cuerpo, 'plain'))
            archivo_adjunto = open(self.ruta_adjunto + self.getFileName(), 'rb')
            adjunto_MIME = MIMEBase('application', 'octet-stream')
            adjunto_MIME.set_payload((archivo_adjunto).read())
            encoders.encode_base64(adjunto_MIME)
            adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % self.getFileName())
            mensaje.attach(adjunto_MIME)
            sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
            sesion_smtp.starttls()
            sesion_smtp.login(remitente, password)
            texto = mensaje.as_string()
            sesion_smtp.sendmail(remitente, destinatarios, texto)
            sesion_smtp.quit()
