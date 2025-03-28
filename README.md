# CoolSpool_email_wrapper
 this allows you to store coolspool command parameters so you can dynamically change the parameters without recompile or create quick automated reports when combined with robot sceduler.
example usage :
we have table sampletbl that contains users. we want to gather a report of the active users and send it to our email. Say you have a query your using to identify this
SELECT * FROM sampletbl WHERE status = 'Active'
we then would setup femailer with the following parameters
![alt text](image-1.png)
then we would call pemailer with the following key parms
Call Pemailer Parm('TESTPGM' 'V1')
and this is the result
3![Screenshot 2025-03-28 at 12 10 18 PM](https://github.com/user-attachments/assets/63adc61e-64fc-46f8-a120-ba8050f44a5d)
