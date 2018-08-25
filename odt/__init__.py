#by Baraza Moses
#email: <clonesoftwares@gmail.com> or <barazayo@gmail.com>


class ODTFile:
    def __init__(self, filename):
        self.filename=filename
        self.libre_doc()

    def libre_doc(self):
        import zipfile
        import xml.etree.ElementTree as ET
        try:
            import os
            if not (os.path.exists(self.filename) and self.filename.endswith('.odt')):
                raise 'File doesnot exists or not a .odt file'
            with zipfile.ZipFile(self.filename, 'r') as stream:
                with stream.open('meta.xml', 'r') as k:
                    xml_string=k.read()
                    tree=ET.fromstring(xml_string)
                    for child in tree:
                        if child.tag == '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}meta':
                            for sub_child in child:
                                if sub_child.tag=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}creation-date':
                                    self.creation_datetime = sub_child.text #parse datetime for file creation
                                elif sub_child.tag=='{http://purl.org/dc/elements/1.1/}date':
                                    self.last_access_datetime = sub_child.text #parse datetime for file access
                                elif sub_child.tag=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}generator':
                                    self.document_generator = sub_child.text #document generator application and environment
                k.close()

                self.text_list=[]

                with stream.open('content.xml', 'r') as content_stream:
                    tree=ET.fromstring(content_stream.read())
                    for child in tree:
                        if child.tag=='{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body': #body tag 
                            for subchild in child:
                                if subchild.tag=='{urn:oasis:names:tc:opendocument:xmlns:office:1.0}text':
                                    for subsubchild in subchild:
                                        if subsubchild.tag=='{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p': #p tag (paragraph tag) containing strings of text
                                            if subsubchild.text==None: #empty tags with no text are empty lines
                                                self.text_list.append('\n') #for empty lines append a new line character
                                            else:
                                                self.text_list.append(subsubchild.text)
                content_stream.close()
                
            stream.close()
        except zipfile.BadZipFile:
            raise 'Read Fail: Bad .odt file'

    def creationdatetime(self):
        return self.creation_datetime

    def accessdatetime(self):
        return self.last_access_datetime

    def documentgenerator(self):
        return self.document_generator

    def readodt(self):
        return self.text_list


# path='/home/coder/Desktop/mydoc.odt'
x=ODTFile('/home/coder/Desktop/mydoc.odt')
print(x.readodt())
print(x.creationdatetime())


