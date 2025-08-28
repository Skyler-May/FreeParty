# 免费安装微软官方 office 2024 专业增强版 全家桶

## 步骤一 [下载部署工具](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

1. 首先在桌面创建一个 `config` 文件夹
2. 打开 `部署工具` 选择 `config` 文件夹将会下载两个文件在此文件中，需要删除 `configuration-Office365-x64` 文件;

## 步骤二 [设置office版本](https://config.office.com/deploymentsettings)

1. 产品和版本
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/1-Products-Versions.png" />

2. 语言
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/2-language.png" />

3. 安装
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/3-install.png" />

4. 更行和升级
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/4-Updates-Upgrades.png" />

5. 授权和激活 [©](https://learn.microsoft.com/zh-cn/deployoffice/vlactivation/gvlks)
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/5-Authorization-Activation.png" />

6. 常规
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/6-conventional.png" />

7. 应用程序首选项
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/7-Application-Preferences.png" />

8. 导出 `XML`  文件到创建好的 `config` 文件夹中
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/8-export.png" />
<img src="https://github.com/Skyler-May/FreeParty/blob/main/office%202024/img/9-export-2.png" />

## 步骤三 （重点）

1. 管理员运行 `cmd`, 并将桌面上的 `config` 路径复制到终端按回车

```bash
# 根据自己的实际路径修改
cd C:\...\...\Desktop\config 

# 下载命令（此过程时间较长，可打开 任务管理器-->性能-->以太网 查看下载状态）
setup /download config.xml

# 安装命令
setup /configure config.xml

# 启动命令 （逐步运行）
cd C:\Program Files\Microsoft Office\Office16 # 如果是安装的32位，需要改成 cd C:\Program Files (x86)\Microsoft Office\Office16
cscript ospp.vbs /sethst:kms.03k.org 
cscript ospp.vbs /act

```
