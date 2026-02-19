@echo off
echo hello here is my test script!
echo color a will change the color to neon green:
color a
echo ver will show windows version:
ver
echo vol will give information about the C drive:
vol
echo hello world printed using a variable:
SET testVar=Hello Wrld
echo %testVar%
pause
echo now we will being to learn about arithmetic operators:
echo set arithmetic operations using the `/a` or `/A` when set-ing a variable
set /A sum=10++10
echo %sum%
echo using the CLI operators are used once in the expression:
echo `set /a sum=10+10`
echo using batch files, operators are used twice in the expressions:
echo `set /a sum=10++10`
echo write comments using `::`
pause
echo in this section we will test input and output
set /p input= Type any input
echo Input is: %input%
pause
echo in this next section we learn about aliases
echo use the doskey keyword to create an aliance for a one time occurance *NOTE - only works in a CLI/PS instance*
echo `doskey docs = cd C:\Users\talop\Documents`
echo `docs`
pause