This code basically display three pictures
in certain order.  

I named this order like "MouseOut" when the
mouse is out of button, "MouseOver" when is
over button and "MouseDown" when mouse button
is pressed.   

To make a button is very easy. It's only you
draw a picture (See the sample in frmHowTo.frm).   

About masks: 
============
In the picture must have a mask,
where is defined the clickable area.
This mask can have two or three colors.
The white color is the invisible area.
The black color is the clickable area.
The third color, (any color, minus black or white)
is the visible area, but not clickable area.
If your picture not have visible and no clickable
area, you must ignore the third color.

About Sounds:  
=============
Now this code have the capability to play  
sounds on click. You can select any sound from  
external file, but the sound must fit into available  
physical memory.

For better performance, it is a good idea to
compile the UserControl like OCX.

It's all.

I'm sending some examples of images too, in
separate zip.  

	Enjoy


	Fausto Cruz Arruda
	cruzarruda@hotmail.com


Sorry for my bad english  :)