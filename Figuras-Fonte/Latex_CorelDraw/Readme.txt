This small script provides the possibilty to support Latex in CorelDraw. 

The file frmLatexEdit.frm contains a simple dialog to create and edit a Latex 
text or equation. The file CorelMacros.bas contain a macros to open the dialog. 
The best way to use the script is to define a hotkey to start the macro 
"GlobalMacros.CorelMacros.LatexEdit".

The script needs latex.exe and dvips.exe (I use MikTex to provide these files). 
These files should be in the path or you need to change the scripts.

Example:

If you want to add a equation in a diagram just open the dialog. Then 
write an equation in latex style, e.g.

$x_1 = \sum_{i=1}^{n} y_i$

and push the Ok-button. The equation will be imported in CorelDraw as curve. If 
you want to change this equation, then select it and start the dialog again. 
There you will be able to edit the equation. The transformation of the new 
equation will be maintained.

I never implemented anything in Visual Basic except this little script. So maybe 
there is a better way to distribute the files or the script can be improved. If 
you know it better, feel free and write me a mail to 

jbender@ira.uka.de

Have fun with it!