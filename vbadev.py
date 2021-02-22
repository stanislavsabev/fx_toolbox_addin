import os 
import re
import xlwings as xw

_addin_name = 'fx_toolbox_v1.4.0.xlam'


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


def set_addin(addin_name: str):
    global _addin_name
    _addin_name = addin_name


def get_addin_book():
    return xw.books[_addin_name]


def import_modules(source_path: str, book_name: str, module_list=None):
    macro = get_addin_book().macro("Import")
    dest_path = os.path.abspath(source_path)
    if module_list:
        macro(source_path, book_name, module_list)
    else:
        macro(source_path, book_name)


def delete_modules(book_name: str, module_list=None):
    macro = get_addin_book().macro("Delete")
    if module_list:
        macro(book_name, module_list)
    else:
        macro(book_name)


def export_modules(dest_path: str, book_name: str, module_list=None):
    macro = get_addin_book().macro("Export")
    dest_path = os.path.abspath(dest_path)
    if module_list:
        macro(dest_path, book_name, module_list)
    else:
        macro(dest_path, book_name)