try:
    import zipfile
    import os
    import xml.etree.ElementTree as ET
except ModuleNotFoundError:
    raise ModuleNotFoundError('Failed to import')


class OdtNotFoundError(Exception):
    def __init__(self, argv):
        self.argv=argv

class BadOdtFileError(Exception):
    def __init__(self, argv):
        self.argv=argv

class OdtFile:
    def __init__(self, filepath):
        self.filepath=filepath
        if not os.path.exists(self.filepath):
            raise OdtNotFoundError('Odt file not found')

    def readodt(self):
        self.linelist=[]
        try:
            with zipfile.ZipFile(odtfile, 'r') as odtstream:
                xmlstring=(odtstream.read('content.xml')).decode(encoding='utf-8')
                tree=ET.fromstring(xmlstring)
                for element in tree.iter():
                    if element.text != None:
                        self.linelist.append(element.text)
            odtstream.close()
            return self.linelist
        except zipfile.BadZipFile:
            raise BadOdtFileError('Odt file is corrupt or broken')

    def odtstat(self):
        with zipfile.ZipFile(odtfile, 'r') as odtstream:
            xmlstring=(odtstream.read('meta.xml')).decode(encoding='utf-8')
            tree=ET.fromstring(xmlstring)
            key=[]
            val=[]
            for element in tree.iter():
                if element.tag=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}document-statistic':
                    for stat_key, stat_value in element.attrib.items():
                        #print(stat_key, '------', stat_value)
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}table-count':
                            key.append('table-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}image-count':
                            key.append('image-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}object-count':
                            key.append('object-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}page-count':
                            key.append('page-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}paragraph-count':
                            key.append('paragraph-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}word-count':
                            key.append('word-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}character-count':
                            key.append('character-count')
                            val.append(int(stat_value))
                        if stat_key=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}non-whitespace-character-count':
                            key.append('non-whitespace-character-count')
                            val.append(int(stat_value))

        odtstream.close()
        return dict(zip(key, val))

    @property
    def odtcreationdate(self):
        with zipfile.ZipFile(odtfile, 'r') as odtstream:
            xmlstring=(odtstream.read('meta.xml')).decode(encoding='utf-8')
            tree=ET.fromstring(xmlstring)
            for element in tree.iter():
                if element.tag=='{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}creation-date':
                    return element.text
        odtstream.close()

x=OdtFile(odtfile)
print(x.readodt())
