# py_grades_list

## Description

This programm aims to send an individual message to all students with their grades contained in an xlsx file according to "grades.xlsx"
This proramm will connect to your email acount and send emails with the following structures ("message_grades.txt"):

`Hello,

  Your grade for X in C is Y.

  So, your final grade F is A.

Regards.`

X, C, Y, F and A will be modified according to the context. 
See f_message Configuration to get more infos. 

## Requirements:

### Python3 packages:

- openpyxl
- ssl
- smtplib
- email.message

### Allow your email server to accept connection from other clients. 

For example in gmail it's by creating an app password that it will work.


## Configuration

All configurations are made in the "CONF VARIABLES"

### sender

It is your email address

### mdp 

It is your password

### host

It is the smptp server of your email.

### notes_file

It is the xlsx file containing all the grades. The structure of this file is given by "grades.xlsx".

### note_ajoutee 

It is the last exam added in the file

### f_message

It is the message sent by the programm with the variables that will be modified according to the context.
"F" will be modified according to if all the exam of the class are filled. 

Not all the exam are filled:

"F" -> ", for now,"

All the exam are filled:

"F" -> ""

"C" is replaced by the class topic (filled in the xlsx file)

"X" is replaced by the name of the exam (filled in the xlsx file).

"Y" is the actual grade of the student in he exam.

"A" is the current final grade of the student in the class.







