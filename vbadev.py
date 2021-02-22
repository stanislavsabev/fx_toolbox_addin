import os 
import re
import xlwings as xw

addin_book: xw.Book = None
addin_name = 'fx_toolbox_v1.4.0.xlsm'


def expand_vba_declarations(vbafile: str):
    vba_types = ['.cls', '.bas']
    if os.path.splitext(vbafile)[-1] not in vba_types:
        raise ValueError(
            'Unsupported file type. Expected one of: {}'.format(', '.join(vba_types)))
    
    patt = re.compile(r'(Public |Private )?(Sub|Function|Property) *.')
    data = []

    with open(vbafile, 'r') as f:
        line = f.readline()

        while line:
            if re.match(patt, line):
                while line.endswith(' _\n'):
                    line = line[:-2]
                    line += f.readline().lstrip()
            data.append(line)
            line = f.readline()
    with open(vbafile, 'w+') as f:
        f.writelines(data)


def import_modules(source_path: str, book_name: str, module_list=None):
    macro = addin_book.macro("Import")
    macro(source_path, book_name, module_list)


def delete_modules(book_name: str, module_list=None):
    macro = addin_book.macro("Delete")
    macro(book_name, module_list)


def export_modules(dest_path: str, book_name: str, module_list=None):
    macro = addin_book.macro("Export")
    macro(dest_path, book_name, module_list)


def set_addin(addin_name: str):
    global addin_book
    addin_book = xw.books[addin_name]


if __name__ == '__main__':
    addin_book: xw.Book = xw.books[addin_name]
