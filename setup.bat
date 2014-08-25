echo "Creating a new directory..."
mkdir %userprofile%\beaverAutoCAD
echo "Copying files to the new directory..."
copy *.* %userprofile%\beaverAutoCAD\
cd %userprofile%\beaverAutoCAD\
echo "Create a shortcut on the desktop"
SHORTCUT.JS
echo "Done installing. Good luck!"