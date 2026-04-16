
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
    print(f"Project Name: {project.getProjectProperties().getProjectTitle()}")
    tasks = project.getTasks()
    print(f"Number of tasks: {tasks.size()}")
    for i in range(min(5, tasks.size())):
        task = tasks.get(i)
        print(f"Task {i}: {task.getName()} (ID: {task.getID()})")
except Exception as e:
    print(f"Error: {e}")
finally:
    jpype.shutdownJVM()
