/*
+----------------------------------------------------------------+
|	WordPress Plugin: Collapsible Comments    								     |
|	http://deuced.net/wpress/collapsible-comments/						     |
|	File Information:																	             |
|	- Collapsible Comments Javascript  												     |
|	- wp-content/plugins/collapsible-comments/ccomments.js         |
| This code is based on Arvind Satyanarayan example which        |
| toggles the visibility of multiple  elements on a page         |
| http://blog.movalog.com/a/javascript-toggle-visibility/        |
+----------------------------------------------------------------+
*/
function xcollapse(id) 
{ 
      var ccommid = document.getElementById(id); 
      if (ccommid != null) {
        if (ccommid.style.display == 'block') 
        { ccommid.style.display = 'none'; }
        else 
        { ccommid.style.display = 'block'; }
                           }
}