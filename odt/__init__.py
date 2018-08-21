#by Baraza Moses
#email: <clonesoftwares@gmail.com> or <barazayo@gmail.com>


class ODTFile:
    def __init__(self, filename):
        self.filename=filename
        self.libre_doc()

    def libre_doc(self):
        try:
            import os
            if not (os.path.exists(self.filename) and self.filename.endswith('.odt')):
                raise 'File doesnot exists or not a .odt file'
            import zipfile
            with zipfile.ZipFile(self.filename, 'r') as stream:
                with stream.open('meta.xml', 'r') as k:
                    xml_string=k.read()
                    import xml.etree.ElementTree as ET
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
            stream.close()
        except zipfile.BadZipFile:
            raise 'Read Fail: Bad .odt file'

    def creationdatetime(self):
        return self.creation_datetime

    def accessdatetime(self):
        return self.last_access_datetime

    def documentgenerator(self):
        return self.document_generator


