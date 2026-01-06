import paramiko
import time
import json
import os
from typing import Dict, List, Optional


class HFSSRemoteSimulator:
    def __init__(self, ssh_config: Dict):
        """
        初始化HFSS远程仿真控制器
        :param ssh_config: SSH连接配置
        """
        self.ssh_config = ssh_config
        self.ssh_client = None
        self.sftp_client = None

    def connect(self) -> bool:
        """建立SSH连接"""
        try:
            self.ssh_client = paramiko.SSHClient()
            self.ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self.ssh_client.connect(
                hostname=self.ssh_config['host'],
                port=self.ssh_config['port'],
                username=self.ssh_config['username'],
                password=self.ssh_config['password']
            )
            self.sftp_client = self.ssh_client.open_sftp()
            print(f"成功连接到 {self.ssh_config['host']}")
            return True
        except Exception as e:
            print(f"SSH连接失败: {str(e)}")
            return False

    def disconnect(self):
        """断开SSH连接"""
        if self.sftp_client:
            self.sftp_client.close()
        if self.ssh_client:
            self.ssh_client.close()
        print("SSH连接已断开")

    def upload_file(self, local_path: str, remote_path: str) -> bool:
        """上传文件到远程主机"""
        try:
            self.sftp_client.put(local_path, remote_path)
            print(f"文件上传成功: {local_path} -> {remote_path}")
            return True
        except Exception as e:
            print(f"文件上传失败: {str(e)}")
            return False

    def execute_command(self, command: str, wait_time: int = 5) -> tuple:
        """执行远程命令"""
        try:
            print(f"执行命令: {command}")
            stdin, stdout, stderr = self.ssh_client.exec_command(command)
            time.sleep(wait_time)

            output = stdout.read().decode('utf-8')
            error = stderr.read().decode('utf-8')

            if error:
                print(f"命令执行错误: {error}")

            return output, error
        except Exception as e:
            print(f"命令执行失败: {str(e)}")
            return "", str(e)

    def setup_hfss_environment(self) -> bool:
        """设置HFSS仿真环境"""
        commands = [
            "mkdir -p D:\\ssh\\tmp\\hfss_simulations",
            "chmod 755 D:\\ssh\\tmp\\hfss_simulations"
            
            "scp -P 22 hostname@IPV4:File_A  File_B"

        ]

        for cmd in commands:
            output, error = self.execute_command(cmd)
            if error:
                return False
        return True




def main():
    # SSH配置
    ssh_config = {
        "host": "172.16.0.1",
        "port": 22,
        "username": "Admin",
        "password": "admin123"
    }

    # 仿真参数
    simulation_params = {
        "solver": "HFSS",
        "cores": "4"
    }

    # 创建仿真控制器
    simulator = HFSSRemoteSimulator(ssh_config)

    try:
        # 连接远程主机
        if not simulator.connect():
            return

        # 设置环境
        if not simulator.setup_hfss_environment():
            print("环境设置失败")
            return

        # 运行仿真 (需要提供实际的HFSS项目文件路径)
        # result = simulator.run_hfss_simulation("path/to/your/project.aedt", simulation_params)
        # print(f"仿真结果: {json.dumps(result, indent=2, ensure_ascii=False)}")

        # 下载结果 (需要先运行仿真)
        # simulator.download_results("./results")

        print("HFSS远程仿真控制器已准备就绪")
        print("请提供有效的HFSS项目文件路径以运行仿真")

    except KeyboardInterrupt:
        print("用户中断操作")
    except Exception as e:
        print(f"发生错误: {str(e)}")
    finally:
        simulator.disconnect()


if __name__ == "__main__":
    main()
