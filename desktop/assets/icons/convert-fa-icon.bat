"c:\Program Files\Inkscape\inkscape.exe" -z -w 16 %1.svg -e %1-16.png
"c:\Program Files\Inkscape\inkscape.exe" -z -w 32 %1.svg -e %1-32.png
magick convert %1-16.png %1-32.png %1.ico