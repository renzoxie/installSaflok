# 最低要求
<ol>
    <li> 管理员权限 </li>
    <li> Windows 8+ / windows Server 2003+ </li>
    <li> PowerShell V5+ </li>
    <li> 如果提示无权运行，需要用管理员权限在PowerShell命令行中输入以下命令并回车<br />set-executionPolicy remoteSigned -force </li>
</ol>

# 可选版本
<ol>
    <li> 5.45 </li>
    <li> 5.68 </li>
</ol>

# 功能
<ol>
    <li> 支持安装在C盘或者D盘 </li>
    <li> 后台共享数据库文件夹 </li>
    <li> 自动安装需要的IIS组件 </li>
    <li> 完成安装后提示软件版本  </li>
    <li> 默认安装PMS测试软件  </li>
    <li> 预制部分配置文件  </li>
</ol>

# 说明
````Powershell
get-help ./installSaflok.ps1 -full
````

# 安装示例
````Powershell
./installSaflok.ps1 -inputDrive 'c' -version '5.45' -property 'Hotel Name' -vendor 'dormakaba'
````

# 版本记录
V2.1 - 加入中文菜单 <br />
V2.0 - 重建、优化架构 <br />
V1.7 - 修复在C盘安装有时报错的问题 <br />
V1.6 - 加入函数 <br />
V1.5 - 加入提示安装版本 <br />
V1.4 - 为5.x版本软件升级SQL Server Express（2008版至2012版）<br />
V1.3 - 加入安装盘验证 <br />
V1.2 - 兼容32位操作系统 <br />
V1.1 - 兼容win7/server-2008R2操作系统 <br />
