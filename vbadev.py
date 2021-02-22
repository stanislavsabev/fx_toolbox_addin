#In[0]
import os 
import re

#In[0]
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

