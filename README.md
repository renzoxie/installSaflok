# Know the mini requirements
<ol>
    <li> Run as administrator </li>
    <li> Windows 8+ / windows Server 2003+ </li>
    <li> PowerShell V5+ </li>
    <li> If system prompt no right to run the script, please run the following command in advance <br />set-executionPolicy remoteSigned -force</li>
    <li> Chocolatey optional </li>
</ol>

# Version Options
<ol>
    <li> 5.45 </li>
    <li> 5.68 </li>
    <li> 6.11 </li>
</ol>

# Features
<ol>
    <li> support install on drive C or drive D </li>
    <li> with destination install drive option </li>
    <li> share folder </li>
    <li> enable IIS features in back end  </li>
    <li> prompt file version after installation  </li>
    <li> with web service tester installed in default  </li>
    <li> with pre-defined config files  </li>
</ol>

# Instruction
````Powershell
get-help ./installSaflok.ps1 -full
````

# Example
````Powershell
./installSaflok.ps1 -inputDrive 'c' -version '5.45' -property 'Hotel Name' -vendor 'dormakaba'
````

# Change log
V2.3 - optimize some codes, add progress bar GUI <br />
V2.2 - compatible with saflok version 6.11, add choco for installation option<br />
V2.1 - add Chinese language<br />
V2.0 - rebuild script structure <br />
V1.7 - fix minor bug for installing on drive c <br />
V1.6 - final release, more functions less lines <br />
V1.5 - add functions, prompt version installed <br />
V1.4 - with SQL Server Express 2012 for version 5.x <br />
V1.3 - drive letter validation <br />
V1.2 - compatible with 32Bit OS <br />
V1.1 - compatible with win7/win2008R2 or lower <br />

# Contributing workflow
Here's how we suggest you go about proposing a change to this project:<br />
<ol>
    <li><a href="https://help.github.com/articles/fork-a-repo/">Fork this project</a> to your account. </li>
    <li><a href="https://help.github.com/articles/creating-and-deleting-branches-within-your-repository">Create a branch</a> for the change you intend to make.</li>
    <li>Make your changes to your fork.</li>
    <li><a href="https://help.github.com/articles/using-pull-requests/">Send a pull request</a> from your fork's branch to our master branch.</li>
</ol>