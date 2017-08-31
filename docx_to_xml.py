import os
import shutil
import xml.etree.ElementTree as et
import zipfile


class ModificaDocx(object):

    scheme = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    def __init__(self, path):
        self.path = path
        self.filename = os.path.basename(path)

    def get_docx_xml(self):
        myzip = zipfile.ZipFile(self.path, 'r')
        xml_content = myzip.read('word/document.xml')
        xml_tree = et.fromstring(xml_content)
        return xml_tree

    def find_text_nodes(self, xml_tree):
        'Iteratore per trovare i nodi testo xml'
        for nodo in xml_tree.iter():
            if nodo.tag == self.scheme + 't':
                yield nodo

    def write_docx(self, xml_tree, filename=''):
        if filename == '':
            filename = self.filename
        else:
            filename = filename + '.docx'

        root_dir = os.getcwd()
        os.makedirs('tmp')
        tmp_dir = os.path.join(root_dir, 'tmp')
        os.chdir(tmp_dir)

        with zipfile.ZipFile(self.path, 'r') as myzip:
            myzip.extractall(tmp_dir)
            myzip.close()

        with open(os.path.join(tmp_dir, 'word/document.xml'), 'w') as f:
            f.write(et.tostring(xml_tree, encoding='unicode'))
            f.close()

        with zipfile.ZipFile(filename, 'w') as f:
            for item in zipfile.ZipFile(self.path, 'r').namelist():
                f.write(os.path.join(tmp_dir, item), item)
            f.close()

        os.rename(os.path.join(tmp_dir, filename), os.path.join(root_dir, filename))
        shutil.rmtree(tmp_dir)


class ModificaDisegni(object):
    'Toolkit per aggiornare il documento per la modifica disegni'

    template = '/home/io/Documents/codes/python/template.docx'

    def __init__(self, author, date, modification_description):
        self.author = author
        self.date = date
        self.modification_description = modification_description

    def new_autohr(self, author):
        self.author = author

    def new_date(self, date):
        self.author = date

    def modification_description(self, modification_description):
        self.author = modification_description

    def fill_template(self, filename):
        template = ModificaDocx(self.template)
        template_xml_tree = template.get_docx_xml()
        for node in template.find_text_nodes(template_xml_tree):
            new_text = 'ciao mamma!!!!'
            node.text = new_text
        template.write_docx(template_xml_tree, filename)


def main():

    mod_1 = ModificaDisegni('io', '18-3-84', 'muhauhuahu')
    mod_1.fill_template('daje')




#    author = input("Mittente :")
#    date = input("Data :")
#    file_name = input("Nome File :")
#
#    mod_1 = ModificaDisegni(file_name, author, date)
#    print(mod_1.file_name)
#
#    mod_1 = ModificaDisegni('daje', 'Pierpaolo', '18/03/1984')
#
#    my_xml_tree = mod_1.get_docx_xml()
#
#    for node in mod_1.find_text_nodes(my_xml_tree):
#        new_text = 'DAJE de PIUUUUU!!!!'
#        node.text = new_text
#
#    mod_1.write_docx(my_xml_tree, 'daje')

if __name__ == '__main__':
    main()

