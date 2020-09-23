DODelete Software for Windows
===============================

Version: 2.0.0
README file updated on: 25 MAY 2006
© 2006 FliptBit Technologies, Inc.


--------------------------
What's new with Version 2
--------------------------

	I have updated the GUI for this program to make it a little more palatable and
	mainstream and a little less weird.

	The main body of the GUI is the TreeView control which is used to drag files and folders
	onto for wiping.  You can drop multiple files and folders onto the control and it will
	recursively read all files and display the file information (path, times, etc.).



---------------------------
*** ABOUT WIPING FILES ***
---------------------------

	I gaurantee that the file(s) you select to wipe CAN NOT BE RECOVERED by any of the
	software recovery programs on the market.  But I can NOT gaurantee that any other
	remnants of the file on your hard drive can not be recovered.  For Example, if you
	have been working on a document for say many months, every time you save it you
	may not be saving the actual file but a copy of it.  The older versions MAY still
	exist on your hard drive.  It is these ghost files that this program DOES NOT wipe.

	Future revisions will however address this problem.

	The patterns used for overwriting are based on Peter Gutmann's paper "Secure Deletion
	of Data from Magnetic and Solid-State Memory" and they are selected to effectively remove
	the magnetic remnants from the hard disk making it impossible to recover the data.

	Other methods include the one defined in the National Industrial Security Program Operating
	Manual of the US Department of Defense and overwriting with pseudorandom data.



----------------
Version History
----------------

	Version 2.0.0 - Created new GUI with drag-n-drop enabled tree view control.
		      - The program will now recursively read folders dropped onto it.
		      - Changed the overwrite patterns to exact DOD specifications.
                      - Added Gutmann wipe method.
		      - Files are now truncated to 0 length to clear reference of previous size.
		      - Program settings are saved in the registry.


	Version 1.0.1 - Added file rename function.  Currently the number of renames is hard-coded to 20.
		      - Added CRC16 function with setable masking to be used with rename function.
		      - Added multiple pass write ability to DOD_elete routine (with minimal speed loss).
		      - Changed Form_Load to only process file and not bother with graphical stuff.
		      - Added UpdateStatusLabel routine to simplify the status messages, and removed the
			DoEvents from it (Stuck with .refresh instead).



--------------
Next Revision
--------------

	- Currently I am working on the code for a registry scanner that will scan all registry keys
	  for traces of the files that have been wiped and the option to delete these keys.  This
	  code is near completion.

	- Will add a routine that scrambles the file dates before deletion.

	- Will add code that will allow the user to wipe the empty space on the hard drive.

	- Will add code that allows for the saving of wipe sessions.	



------------------
Program Operation
------------------

  I used the NISPOM DOD 5220.22-M standard as a guideline to write the
  code.  My old version of the program used the Print statement to
  write to the file but as I began learning about C++ I learned about
  using API in Visual Basic.  So now the program uses the CreateFile
  and WriteFile API, as well as the FlushFileBuffers API to try and
  ensure that any buffered data gets written to disk.

  The program is easy to operate, just drag and drop any files or folders
  onto the list and click the Wipe control.  The program will then start
  wiping the files as per the settings you select.

  NOTE: You can stop the wiping at any time.  When you press the stop button
        the program stops wiping AFTER it has completed with its CURRENT overwrite
        operation.  Any additional files in the list will remain, but the current
        it was operating on WILL BE LOST!!


-------------
Known Issues
-------------

	* Everytime a drag-and-drop occurs, any previous data is removed and only the
	  new drag-and-dropped files show up.  The list does not currently get appended.

	* Severe lack of command line functions -- I know -- working on it.

	* When dragging a file and an empty folder onto the app, the program displays that
	  the folder is empty and does NOT add the valid file to the wipe list.

	* The FILE_FLAG_NO_BUFFERING flag is not used yet because I still have to
          write the code to sector-align the data array before the write.  The code
	  is still sound without this flag, but having this flag set would further
	  the security of the program.  The code in C would look something like this:


  		char buf[2 * SECTOR_SIZE - 1], *p;

  		p = (char *) ((DWORD) (buf + SECTOR_SIZE - 1) & ~(SECTOR_SIZE - 1));
  		h = CreateFile(argv[1], GENERIC_READ | GENERIC_WRITE,
      		FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, CREATE_ALWAYS,
      		FILE_ATTRIBUTE_NORMAL | FILE_FLAG_NO_BUFFERING, NULL);
  		WriteFile(h, p, SECTOR_SIZE, &dwWritten, NULL);


	* Appropriate error handling has not been implemented.  There is some, but for
	  this program to be public release, it needs to be idiot-proofed a liitle more.

	* GUI is getting better


-----------
Disclaimer
-----------

  TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, IN NO EVENT SHALL
  FLIPTBIT TECHNOLOGIES BE LIABLE FOR ANY SPECIAL, INCIDENTAL, INDIRECT,
  PUNITIVE OR CONSEQUENTIAL DAMAGES WHATSOEVER (INCLUDING, BUT NOT LIMITED
  TO, DAMAGES FOR:  LOSS OF PROFITS, LOSS OF CONFIDENTIAL OR OTHER INFORMATION,
  BUSINESS INTERRUPTION, PERSONAL INJURY, LOSS OF PRIVACY, FAILURE TO MEET
  ANY DUTY (INCLUDING OF GOOD FAITH OR OF REASONABLE CARE), NEGLIGENCE, AND
  ANY OTHER PECUNIARY OR OTHER LOSS WHATSOEVER) ARISING OUT OF OR IN ANY
  WAY RELATED TO THE USE OF OR INABILITY TO USE THE OS COMPONENTS OR THE
  SUPPORT SERVICES, OR THE PROVISION OF OR FAILURE TO PROVIDE SUPPORT SERVICES,
  OR OTHERWISE UNDER OR IN CONNECTION WITH ANY PROVISION OF THIS SUPPLEMENTAL
  EULA, EVEN IF FLIPTBIT TECHNOLOGIES OR ANY SUPPLIER HAS BEEN ADVISED OF THE
  POSSIBILITY OF SUCH DAMAGES. 


---------
Liscense
---------

  Copyright (C) 2006  John R. Reid IV (aka FliptBit)
  
  This program is free software; you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation; either version 2 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program; if not, write to the Free Software
  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA



© 2006 FliptBit Technologies, Inc.