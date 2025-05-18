import os
import shutil
import win32com.client
import pythoncom
import win32file
import win32con
import pywintypes # For pywintypes.Time()
import time       # 引入 time 模块用于延迟

def set_file_times(target_path, source_path):
    max_retries = 5
    retry_delay = 0.5 # 秒

    for attempt in range(max_retries):
        try:
            source_stat = os.stat(source_path)
            win_creation_time = pywintypes.Time(source_stat.st_ctime)
            win_access_time = pywintypes.Time(source_stat.st_atime)
            win_modify_time = pywintypes.Time(source_stat.st_mtime)

            handle = win32file.CreateFile(
                target_path,
                win32con.GENERIC_WRITE,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None
            )
            win32file.SetFileTime(handle, win_creation_time, win_access_time, win_modify_time)
            win32file.CloseHandle(handle)
            print(f"成功将源文件 {os.path.basename(source_path)} 的时间戳应用到 {os.path.basename(target_path)}")
            return # 成功后退出函数
        except pywintypes.error as e: # 更具体地捕获 pywin32 的错误
            # winerror 32 is ERROR_SHARING_VIOLATION
            if e.winerror == 32 and attempt < max_retries - 1:
                print(f"警告: 文件 {os.path.basename(target_path)} 被占用，将在 {retry_delay} 秒后重试设置时间戳 (尝试 {attempt + 2}/{max_retries})...") # attempt is 0-indexed
                time.sleep(retry_delay)
            else:
                print(f"警告: 无法为文件 {os.path.basename(target_path)} 设置与源文件 {os.path.basename(source_path)} 相同的时间戳: {e}")
                return # 达到最大重试次数或遇到其他 pywin32 错误
        except FileNotFoundError:
             print(f"警告: 无法找到文件 {target_path} 或 {source_path} 来设置时间戳。")
             return
        except Exception as e: # 捕获其他可能的异常
            print(f"警告: 设置文件 {os.path.basename(target_path)} 时间戳时发生意外错误: {e}")
            return
    # 如果循环结束仍未成功
    print(f"警告: 多次尝试后仍无法为文件 {os.path.basename(target_path)} 设置时间戳，文件可能持续被占用。")


def create_old_files_folder(source_directory):
    old_files_folder_name = "旧格式文件"
    old_files_path = os.path.join(source_directory, old_files_folder_name)
    if not os.path.exists(old_files_path):
        try:
            os.makedirs(old_files_path)
            print(f"创建文件夹: {old_files_path}")
        except OSError as e:
            print(f"错误：无法创建文件夹 '{old_files_path}': {e}")
            return None
    else:
        print(f"文件夹 '{old_files_path}' 已存在.")
    return old_files_path

def convert_doc_to_docx(source_directory, old_files_path):
    if old_files_path is None:
        return

    print("开始处理 DOC 文件...")
    word_app = None
    # Initialize COM for this thread, ensure it's done once per thread using it
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0 # wdAlertsNone = 0

        for root, _, files in os.walk(source_directory):
            normalized_root = os.path.normpath(root)
            normalized_old_files_path = os.path.normpath(old_files_path)
            if normalized_root == normalized_old_files_path or normalized_root.startswith(normalized_old_files_path + os.sep):
                # print(f"跳过已归档目录 (DOC): {root}") # Can be verbose, uncomment if needed
                continue
            
            for file in files:
                if file.lower().endswith(".doc") and not file.lower().startswith("~"):
                    doc_file_path = os.path.join(root, file)
                    docx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".docx")
                    
                    print(f"尝试处理 {doc_file_path} ...")
                    doc = None 
                    should_move_original = False

                    try:
                        if os.path.exists(docx_file_path):
                            print(f"警告: 目标文件 {docx_file_path} 已存在。跳过转换。")
                            should_move_original = True 
                        else:
                            print(f"正在转换 {doc_file_path} 为 {docx_file_path} ...")
                            doc = word_app.Documents.Open(doc_file_path, ReadOnly=True, PasswordDocument="")
                            doc.SaveAs2(docx_file_path, FileFormat=12) # wdFormatXMLDocument = 12
                            doc.Close(SaveChanges=0) # Close document immediately after saving
                            doc = None # Ensure doc object is cleared
                            print(f"转换成功: {docx_file_path}")
                            set_file_times(docx_file_path, doc_file_path)
                            should_move_original = True
                        
                    except pythoncom.com_error as ce:
                        error_message = str(ce).lower()
                        hresult = getattr(ce, 'hresult', 0)
                        password_keywords = ["password", "密码", "protected", "-2146824422", "-2146822422", "incorrect password", "incorrect document password"]
                        office_problem_keywords = ["office检测到此文件存在一个问题", "office has detected a problem with this file"]
                        
                        is_password_error_by_msg = any(keyword in error_message for keyword in password_keywords)
                        is_office_problem = any(keyword in error_message for keyword in office_problem_keywords)
                        is_password_error_by_code = hresult == -2146824422 or hresult == -2146822422
                        
                        if is_password_error_by_msg or is_password_error_by_code:
                            print(f"文件 {doc_file_path} 受密码保护或打开时需要密码，跳过转换。错误: {ce}。原始文件将保留在原位。")
                        elif is_office_problem or hresult == -2147352567: # General COM error often related to file issues
                             print(f"文件 {doc_file_path} Office检测到问题或无法打开，跳过转换。错误: {ce}。原始文件将保留在原位。")
                        else:
                            print(f"处理文件 {doc_file_path} 时发生COM错误: {ce}。原始文件将保留在原位。")
                    except Exception as e:
                        print(f"处理文件 {doc_file_path} 失败: {e}。原始文件将保留在原位。")
                    finally:
                        if doc: # If an error occurred after Open but before Close
                            try:
                                doc.Close(SaveChanges=0)
                            except Exception as e_close:
                                print(f"关闭文档 {os.path.basename(doc_file_path)} 时额外出错: {e_close}")
                        
                        if should_move_original:
                            # Ensure the original file still exists before attempting to move
                            if os.path.exists(doc_file_path):
                                try:
                                    print(f"正在移动 {doc_file_path} 到 {old_files_path}...")
                                    shutil.move(doc_file_path, os.path.join(old_files_path, file))
                                    print(f"移动成功: {doc_file_path} -> {old_files_path}")
                                except Exception as e_move:
                                    print(f"移动文件 {doc_file_path} 失败: {e_move}")
                            else:
                                print(f"警告: 原始文件 {doc_file_path} 在尝试移动前已不存在。")
    except pythoncom.com_error as e_dispatch:
        print(f"初始化Word Dispatch或Word全局操作时出错: {e_dispatch}")
    except Exception as e:
        print(f"初始化Word或处理DOC文件时发生未知错误: {e}")
    finally:
        if word_app:
            try:
                word_app.Quit(SaveChanges=0)
            except Exception as e_quit:
                print(f"尝试退出Word时发生错误: {e_quit}")
        pythoncom.CoUninitialize()

def convert_xls_to_xlsx(source_directory, old_files_path):
    if old_files_path is None:
        return

    print("开始处理 XLS 文件...")
    excel_app = None
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        excel_app = win32com.client.Dispatch("Excel.Application")
        try:
            excel_app.DisplayAlerts = False
        except pythoncom.com_error as e_alerts:
            print(f"警告: 无法设置Excel的DisplayAlerts属性: {e_alerts}. 尝试继续...")
        # excel_app.Visible = True # 调试时可以设为True

        for root, _, files in os.walk(source_directory):
            normalized_root = os.path.normpath(root)
            normalized_old_files_path = os.path.normpath(old_files_path)
            if normalized_root == normalized_old_files_path or normalized_root.startswith(normalized_old_files_path + os.sep):
                # print(f"跳过已归档目录 (XLS): {root}") # Can be verbose
                continue

            for file in files:
                if file.lower().endswith(".xls") and not file.lower().startswith("~"):
                    xls_file_path = os.path.join(root, file)
                    xlsx_file_path = os.path.join(root, os.path.splitext(file)[0] + ".xlsx")

                    print(f"尝试处理 {xls_file_path} ...")
                    workbook = None
                    should_move_original = False
                    try:
                        if os.path.exists(xlsx_file_path):
                            print(f"警告: 目标文件 {xlsx_file_path} 已存在。跳过转换。")
                            should_move_original = True
                        else:
                            print(f"正在转换 {xls_file_path} 为 {xlsx_file_path} ...")
                            workbook = excel_app.Workbooks.Open(
                                xls_file_path,
                                UpdateLinks=0,
                                ReadOnly=True,
                                Format=None,
                                Password="",
                                IgnoreReadOnlyRecommended=True,
                                CorruptLoad=1 # xlRepairFile
                            )
                            
                            if hasattr(workbook, 'HasPassword') and workbook.HasPassword:
                                print(f"文件 {xls_file_path} 打开后仍指示受密码保护，跳过转换。原始文件将保留在原位。")
                                # Ensure workbook is closed if opened but password protected
                                if workbook: workbook.Close(SaveChanges=False)
                                workbook = None
                            else:
                                workbook.SaveAs(xlsx_file_path, FileFormat=51) # xlOpenXMLWorkbook = 51
                                workbook.Close(SaveChanges=False) # Close workbook immediately
                                workbook = None # Ensure workbook object is cleared
                                print(f"转换成功: {xlsx_file_path}")
                                set_file_times(xlsx_file_path, xls_file_path)
                                should_move_original = True
                        
                    except pythoncom.com_error as ce:
                        error_message = str(ce).lower()
                        hresult = getattr(ce, 'hresult', 0)
                        password_keywords = ["password", "密码", "protected", "cannot open the specified file"]
                        is_password_error_by_msg = any(keyword in error_message for keyword in password_keywords)
                        is_password_error_by_code = hresult == -2146827284 # Excel password error
                        
                        if is_password_error_by_msg or is_password_error_by_code:
                            print(f"文件 {xls_file_path} 打开/保存时遇到问题（可能被识别为密码保护或格式不受支持），跳过转换。错误: {ce}。原始文件将保留在原位。")
                        elif hresult == -2147352567: # General COM error often related to file issues
                             print(f"文件 {xls_file_path} Office检测到问题或无法打开，跳过转换。错误: {ce}。原始文件将保留在原位。")
                        else:
                            print(f"处理文件 {xls_file_path} 时发生COM错误: {ce}。原始文件将保留在原位。")
                    except Exception as e:
                        print(f"处理文件 {xls_file_path} 失败: {e}。原始文件将保留在原位。")
                    finally:
                        if workbook: # If an error occurred after Open but before Close
                            try:
                                workbook.Close(SaveChanges=False)
                            except Exception as e_close:
                                print(f"关闭工作簿 {os.path.basename(xls_file_path)} 时额外出错: {e_close}")
                        
                        if should_move_original:
                            if os.path.exists(xls_file_path):
                                try:
                                    print(f"正在移动 {xls_file_path} 到 {old_files_path}...")
                                    shutil.move(xls_file_path, os.path.join(old_files_path, file))
                                    print(f"移动成功: {xls_file_path} -> {old_files_path}")
                                except Exception as e_move:
                                    print(f"移动文件 {xls_file_path} 失败: {e_move}")
                            else:
                                print(f"警告: 原始文件 {xls_file_path} 在尝试移动前已不存在。")

    except pythoncom.com_error as e_dispatch:
        print(f"初始化Excel Dispatch或Excel全局操作时出错: {e_dispatch}")
    except Exception as e:
        print(f"初始化Excel或处理XLS文件时发生未知错误: {e}")
    finally:
        if excel_app:
            try:
                excel_app.Quit()
            except Exception as e_quit:
                print(f"尝试退出Excel时发生错误: {e_quit}")
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    source_dir = os.path.dirname(os.path.abspath(__file__))
    # 或者，如果您想让用户输入目录：
    # source_dir = input("请输入要处理的文件夹路径: ")
    # if not os.path.isdir(source_dir):
    #     print(f"错误：提供的路径 '{source_dir}' 不是一个有效的文件夹。")
    #     exit()

    print(f"将在以下目录执行操作: {source_dir}")
    
    old_files_destination = create_old_files_folder(source_dir)
    
    if old_files_destination:
        convert_doc_to_docx(source_dir, old_files_destination)
        convert_xls_to_xlsx(source_dir, old_files_destination)

    print("文件转换和移动操作完成。")
    print("请注意：此脚本依赖 pywin32 库。如果尚未安装，请运行 'pip install pywin32' 进行安装。")
    print("加密文件、Office检测到问题的文件或无法正确打开的文件将被跳过，并保留在原位。")
    print("设置文件时间戳时，如果文件被占用，脚本会尝试几次。")
