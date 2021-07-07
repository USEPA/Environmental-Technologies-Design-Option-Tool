for %%f in (*.*) do call check.bat %%f
xcopy a:\*.zip .
pkunzip -o *.zip .
xcopy *.zip bak
del *.zip


