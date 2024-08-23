<h1>Installation notes</h1>

<p>The quickest way is to use the Yeoman Generator for Office Add-Ins, which requires the latest LTS version of Node.js.</p>

````
npm install -g yo generator-office
````

<p>This line creates a project.</p>

````
yo office
````

<p>When prompted, provide the following information to create your add-in project.</p>
<ul>
<li>Choose a project type: Office Add-in Task Pane project</li>
<li>Choose a script type: Javascript</li>
<li>What do you want to name your add-in? My Office Add-in</li>
<li>Which Office client application would you like to support? Excel</li>
</ul>

<p>In the folder where you created your project, replace these two files with the ones in this repository.</p>

````
/src/taskpane/taskpane.js
/src/taskpane/taskpane.html
````
