import os
import shutil
import json
import sys
import getopt
import time

sys.path.append(r"F:\下载\备份数据\python\PythonProgram\MyLib")

#pyisntallerpath:pyinstaller.py绝对路径，r"E:\\python_64\\pyinstaller-develop\\pyinstaller.py"
# arg : -F 生成一个exe，-w，取消控制台 -i 更改exe图标
#icopath ： ico图标绝对路径
#pypath ：被打包的python文件绝对路径
#filenewname : 生成的exe文件的名字
#file_to :在哪儿生成文件
class PyIstaller:
    def __init__(self,pyinstallerpath=None,arg = "-F -w -i",icopath=None,pypath=None,file_to=None,program_name="my-program"):
        '''
        :parameter
        '''
        # self.path = os.getcwd()
        self.file_version = None
        self.update_data =None
        self.pyinstaller_path = pyinstallerpath
        self.ico_path = icopath
        self.py_path = pypath
        self.arg = arg
        self.py_name = program_name
        if os.path.isdir(file_to) == False:
            os.mkdir(file_to)
        self.file_to = file_to +os.sep + self.py_name
        self.base_path = os.path.split(self.py_path)[0]
        self.exe_name = "run.exe"
        

    def run(self):
        # path =self.file_to + os.sep                      #本程序目录,因为生成exe在本程序的路径下
        build_path = self.base_path + os.sep + 'build'         #生成的build文件路径
        print("build 目录为：{build_path}".format(build_path=build_path))
        dist_path = self.base_path + os.sep + 'dist'           #生成的dist文件路径，生成的exe存在这里
        print("dist 目录为:{dist_path}".format(dist_path=dist_path))
        exe_path = self.base_path + os.sep + 'dist' + os.sep + self.exe_name  # 生成的exe的路径
        print("exe 目录为:{exe_path}".format(exe_path = exe_path))
        new_exe_path = self.file_to + os.sep + self.exe_name
        if os.path.isdir(build_path) is True:            #预移除可能存在的上一次生成的build和dist路径
            shutil.rmtree(build_path)
            print("预移除{}".format(build_path))
        if os.path.isdir(dist_path) is True:
            shutil.rmtree(dist_path)
            print("预移除{}".format(dist_path))
        command = self.pyinstaller_path + " " + self.arg + " " + self.ico_path + " " + self.py_path  # 编辑指令
        #执行指令，知道成功生成EXE文件
        print("准备执行指令{}".format(command))    #执行指令
        os.system(command)
        print("创建成功")#生成成功
        if "-F" in self.arg:  # 如果执行的是-F指令
            time.sleep(1)
            if os.path.isfile(new_exe_path) is True:       #移除可能存在的目标文件夹的exe文件
                os.remove(new_exe_path)
                print("移除原exe文件{}".format(new_exe_path))
            if os.path.isdir(build_path) is True:          #移除生成exe过程中产生的build文件
                shutil.rmtree(build_path)
                print("移除bild文件{}".format(build_path))
            time.sleep(1)
            if os.path.isfile(exe_path) is True:        #如果dist内exe文件存在
                if os.path.isdir(self.file_to) is False:
                    os.mkdir(self.file_to)
                    print("创建目标文件")
                shutil.move(exe_path, new_exe_path)             #将文件从当前生成目录移到目标目录
                print("加载{}".format(new_exe_path))
            time.sleep(1)
            if os.path.isdir(dist_path) is True:       #删除生成的dist文件
                shutil.rmtree(dist_path)
                print("移除{}".format(dist_path))
        else:
            exe_path = self.base_path + os.sep + 'dist' + os.sep + self.exe_name.split(".")[0]
            # new_exe_path = self.file_to+ os.sep+ exe_name.split(".")[0]
            print(new_exe_path)
            time.sleep(2)
            if os.path.isdir(build_path) is True:          #如果不是-F，则移除全部内容...
                shutil.rmtree(build_path)
                print("移除bild文件{}".format(build_path))
            if os.path.isdir(self.file_to) is True:          #如果不是-F，则移除全部内容...
                shutil.rmtree(self.file_to)
                print("移除原安装文件{}".format(self.file_to))
            shutil.move(exe_path, self.file_to)  # 将文件从当前生成目录移到目标目录
            if os.path.isdir(dist_path) is True:       #删除生成的dist文件
                shutil.rmtree(dist_path)
                print("移除{}".format(dist_path))
        print("exe文件生成成功")
        # self.__build_config_file()
        print("完成软件创建")
        self._copy_static()
        print("完成static文件复制")

    def _copy_static(self):
        '''
        :parameter
        '''
        static_path = self.base_path + os.sep + r"static" # 旧的 static 目录
        new_static_path = self.file_to+ os.sep + r"static" # 新的 static 目录
        # if os.path.isabs(copytree):
        try:
            shutil.copytree(static_path,new_static_path)
        except Exception as e:
            print(e)
        
def usage():
    print("BHP NET TOOL")
    print("")
    print("-F                    --Single File")
    print("-w                    --No Windows")
    print("-i                    --add the exe ico")
    print("-fp                   --py path")
    print("-ep                   --exe path")
    print("-u                    --pyinstallerpath")
    sys.exit(0)

if __name__ == "__main__":
    arg = ""
    SingleFile =False
    NoWindow =False
    IconUse =False
    PyPath = False
    ExePath=False
    command = False
    path_pyinstaller = "pyinstaller"
    file_to = None
    path_py = None
    program_name = None
    
    if not len(sys.argv[1:]):
        print("请输入必要指令！！！")
        usage()

    try:
        opts,args = getopt.getopt(sys.argv[1:],"hFwif:e:u:c",["help","singlefile","nowindow","iconuse","pypath","exepath","pyinstallerpath","command"])
    except getopt.GetoptError as err:
        print(str(err))
        usage()
        sys.exit(0)
        
    for o,a in opts:
        if o in ("-h","--help"):
            usage()
        elif o in ("-F","--singlefile"):   #生成单一文件
            SingleFile = True
        elif o in ("-w","--nowindow"):
            NoWindow = True
        elif o in ("-i","--iconuse"):
            IconUse = True
        elif o in ("-f","--pypath"):
            PyPath = True
            path_py = a
        elif o in ("-e","--exepath"):
            ExePath = True
            file_to = a
        elif o in ("-c","--command"):
            command = True
        elif o in("-u","pyinstallerpath"):
            path_pyinstaller = a
        else:
            print("Unhandled Option:",o)
            assert False
    
    if PyPath is False or ExePath is False or path_py is None or not os.path.isfile(path_py):
        print("必须输入python文件绝对路径（且python文件名字必须为run.py），以及pyinstaller文件绝对路径")
        usage()
        sys.exit(0)
    if path_py.split(os.sep)[-1] != "run.py":
        print("运行文件名必须为run.py")
        usage()
        sys.exit(0)
    
    try:
        program_name = path_py.split(os.sep)[-2]
    except Exception as e:
        program_name = "my-program"
    
    program_name = program_name if program_name != None else "my-program"
    
    if SingleFile:
        arg +=" -F "
    if NoWindow:
        arg +=" -w "
    if IconUse:
        arg +=" -i "
    if command:
        arg += " -c"
    
    base_path, py_name = os.path.split(path_py)
    path_ico = base_path + os.sep + "logo.ico"

    print("SingleFile:{},NoWindow:{},IconUse:{},PyPath:{},ExePath:{}".format(SingleFile, NoWindow, IconUse, PyPath,
                                                                             ExePath))
    if os.path.isfile(path_ico) == False:
        print("未指定ICO路径,使用默认ICO...")
        path_ico = r"E:\\OneDrive\\python\\PythonExe\\normal.ico"
    
    pyinstaller = PyIstaller(pyinstallerpath=path_pyinstaller, arg=arg, icopath=path_ico, pypath=path_py,
                              file_to=file_to,program_name = program_name)
    pyinstaller.run()