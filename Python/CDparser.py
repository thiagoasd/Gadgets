from os import listdir;
from os.path import isdir, join;
import re, sys;

path = str(sys.argv[1]);
print path;

for f in listdir(path):
    cpath = join(path, f);
    print cpath;
    if isdir(cpath):
        print f + "lol";
        m = re.search('(\[.*\])(.*)(\[.*\])', cpath);
        print(m.group(0));
