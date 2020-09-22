I found this scheme somewhere on the internet, while searching for some ideas to
use to make a registerable shareware product out of a program I had written.  It
uses a pretty simple algorithm to generate a password, so I changed the algorithm
to make the password a bit more difficult to break.  I also added the disk serial
number into the algorithm, so that the password would only be usable on a single
computer.  (You can modify your program very easily to eliminate this feature, if
you want to release a program with a site license.  It's still as secure, but the
password depends on a manual seeding of the random number generator, so you'll have
to keep a record of the few site-licensed programs you sell [we should all have
this problem].)

The algorithm is, basically, very simple.  First, we get the serial number of drive
C:.  We then use this number to seed the random number generator (after doing a few
arithmetic manipulations - feel free to change the numbers).  (If anyone can come up
with a better algorithm, something so that any change in a single character in either
the disk serial number or the name will change the random seed, remembering that the
seed must be a positive integer, please let me know.)

We then fill an array of 25 characters (we'll accept a maximum of 25
characters in the name) with random numbers.  This can, of course be expanded.

To create the registration code, we exclusive OR (Xor) each character in the name
with the corresponding character in the array.

We can set the length of the registration code to anything up to 25 characters, so
choose how much work you want to give people.  The default is the length of the name
or 10, whichever is longer.

If you need a manually-set code (for site licenses, for example), just change the
manipulations with the disk serial number to a manually input number.  Limit it to
positive int, that is, 32767.

The sample project is an example of how to use the registration number to limit what
the application can do if the program is not registered.  It's really trivial, in that
it simply sets the caption of a label to tell you whether the program is registered or
not, but you can use the global variable, NotRegistered, to make your program do whatever
you want for the true (-1) and false (0) cases.

If you want to email me, my address is aklein@villagenet.com, but don't expect a response
in 5 minutes - I read my email about once a day.  Most days.

Al Klein
