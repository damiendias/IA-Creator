# IA-Creator

Creates an Episerver scheduled job that looks in the app_data folder for an excel spreadsheet named content.xlsx. If found the scheduled job processes the spreadsheet importing pages into the CMS under a designated parent identifier. The parent identifier is set within the appSettings section in web.config. An example key is below:

\<add key="ParentPageId" value="693" \/\>

The format of the spreadsheet needs to be as follows:

Column A = Content type Name<br />
Column B = page name<br />
Column C = page tree level (i.e 0 = first level under designated parent page, 1 = second level etc.)<br />
Column D+ = One property per column and you can fill in as many columns as you like. The format for the column is [PropertyName]:[Value]. So an example is like:

MainBody:This is the content for the main body

Sample spreadsheet is attached in the Repo...
