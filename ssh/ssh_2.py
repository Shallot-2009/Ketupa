import paramiko
import pyaedt
import pythoncom
import win32com.client
import time


ssh = paramiko.SSHClient()           # 允许连接不在know_hosts文件中的主机
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

def deploy_and_run(host, port, user, pwd, local_path, remote_path):
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        # 建立SSH连接
        ssh.connect(hostname=host,
                    port=port,
                    username=user,
                    password=pwd)

        # 执行命令
        stdin, stdout, stderr = ssh.exec_command(f"Hostname")
        print(stdout.read().decode())

    except paramiko.AuthenticationException:
        print("认证失败，请检查用户名和密码")
    except paramiko.SSHException as e:
        print(f"SSH连接错误: {str(e)}")
    except Exception as e:
        print(f"其他错误: {str(e)}")
    finally:
        ssh.close()


if __name__ == '__main__':
    deploy_and_run(host='172.16.0.1',
                   port=22,
                   user='Admin',
                   pwd='admin123',
                   local_path='local_script.py',
                   remote_path='/home/user/remote_script.py')



# 初始化COM接口
hfss = win32com.client.Dispatch("AnsoftHfss.HfssScriptInterface")

# 连接到远程服务器
hfss.ConnectToServer("172.16.0.71", "AnsoftHfss.HfssScriptInterface")
