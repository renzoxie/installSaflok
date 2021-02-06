>
> Compatible with win8*,win10*, win2012*,win2016*, have not be tested on win7 yet<br />
> For win2008R2, need Powershell version 5 or above<br />


# features
* support install on drive C or drive D
* with destination install drive option 
* share folder in back end  
* enable IIS features in back end  
* prompt file version after installation  
* with web service tester installed in default  
* with pre-defined config files  

# instruction
````Powershell
get-help ./installSaflok.ps1 -full
````

# example
````Powershell
./installSaflok.ps1 -inputDrive 'c' -version '5.45' -property 'Hotel Name' -vendor 'dormakaba'
````

# change log
V2.0 - rebuild script structure <br />
V1.7 - fix minor bug for installing on drive c <br />
V1.6 - final release, more functions less lines <br />
V1.5 - add functions, prompt version installed <br />
V1.4 - with SQL Server Express 2012 for version 5.x <br />
V1.3 - drive letter validation <br />
V1.2 - compatible with 32Bit OS <br />
V1.1 - compatible with win7/win2008R2 or lower <br />

# contributing workflow
Here's how we suggest you go about proposing a change to this project:<br />
<ol>
  <li><a href="https://help.github.com/articles/fork-a-repo/">Fork this project</a> to your account. </li>
  <li><a href="https://help.github.com/articles/creating-and-deleting-branches-within-your-repository">Create a branch</a> for the change you intend to make.</li>
  <li>Make your changes to your fork.</li>
    <li><a href="https://help.github.com/articles/using-pull-requests/">Send a pull request</a> from your fork's branch to our master branch.</li>
</ol>
