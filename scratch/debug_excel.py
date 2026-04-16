
import os, sys
_JAVA_HOME = os.path.expanduser("~/jdk21/Contents/Home")
os.environ["JAVA_HOME"] = _JAVA_HOME
_jvm_lib = os.path.join(_JAVA_HOME, "lib", "server", "libjvm.dylib")
os.environ.setdefault("JPY_JVM", _jvm_lib)

import jpype
if not jpype.isJVMStarted():
    jpype.startJVM()

from net.sf.mpxj.reader.UniversalProjectReader import UniversalProjectReader

excel_file = "/Users/ahmedfawzy/Documents/MS project/TLD F&O and RED365 Project Plan.xlsx"
reader = UniversalProjectReader()
try:
    project = reader.read(excel_file)
    print("Columns found in project:")
    # MPXJ maps Excel columns to internal fields.
    # Let's see what tasks look like.
    tasks = project.getTasks()
    if tasks.size() > 0:
        first_task = tasks.get(1) # Index 0 is often summary
        print(f"Name: {first_task.getName()}")
        print(f"Notes: {first_task.getNotes()}")
except Exception as e:
    print(f"Error: {e}")
finally:
    jpype.shutdownJVM()
