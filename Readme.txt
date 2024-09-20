=====================================
http://www.novell.com/coolsolutions/tools/14871.html
=====================================

Sometimes you want to run or distribute a program,
and want to be able to hide the window(s) it creates.

Introducing RunHide.exe,
an utility program that enables you to do just that.

RunHide /?


SYNTAX: RunHide (program) (parameters) (TIMEOUT=seconds) (TERM)

Run any named program as a hidden window.
Programs launched with RunHide should not require any user interaction.
Should the hidden program require any user interaction to continue, 
or pauses due to an error, the program will remain in memory until 
the user logs out of the computer or the computer is rebooted.

To accomodate any potential errors like this,
specify a TIMEOUT value in seconds. 
If after the specified number of seconds,
the launched program has not terminated, it will be forcefully terminated.

If you did not specify a TIMEOUT value,
but you find the program is still running hidden,
you may terminate the program by specifying the TERM option.

For example:

      RunHide MyProgram TERM

This will kill all instances of MYPROGRAM.EXE

=====================================
http://www.extendingflash.com/utilities/runhide.html
=====================================


Run a program in a hidden window
   
When you EXEC a batch file, a DOS command window is opened and is either fully visible
or will be relegated to the task bar. In either case the effect detracts from the 
quality of your presentation. RUNHIDE fixes this cosmetic problem by hiding the window
 of the program you are launching. 

usage: runhide.exe program

-This actionScript command fscommand("EXEC", "runhide.exe" + chr(9) + "hideme.bat"); 
will launch hideme.bat but the normal DOS window will be hidden from view. 
The BAT file may open other windows and they will appear normally. 


-If you use the old Flash trick of using tabs in place of spaces to pass parameters
to program, RUNHIDE may not work properly. It depends on whether the application
being launched can handle tabs on its command line. This SHOULD work okay with BAT files,
but EXE files have to be tested on a case by case basis. 

-Flash MX does not allow spaces or tabs in the argument to the EXEC command.
This means that the RUNHIDE utility can't be called directly from Flash MX
to pass parameters to the program. You would have to create an EXE wrapper
to launch hidden windows in Flash MX. 


-In a batch file the you can pass whatever you need on the command line for program
which provides a workaround for Flash 5 limitations. 


-If you're working with SWF Studio you can use spaces and pass parameters to program
but it's easier just to use the RUN and RUNMODE actionScript fscommands to accomplish the same task.  
