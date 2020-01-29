pkzip -x*.zip -x*.mdb %1.zip *.*
xcopy %1.zip bak
del a:\*.*
xcopy %1.zip a:\
del %1.zip


