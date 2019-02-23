ReSharper.AutoFormatOnSave
==========================

Keeping code formatted the easy way
-----------------------------------

If you’re like me, you prefer to have your code tidy, and being a lazy guy, you have defined a ReSharper Code Cleanup template. The problem however, is that you might forget to run the Code Cleanup command every time you change the file. So I decided to go ahead and automate that by building a Visual Studio extension which will run the Silent Code Cleanup command every time a file in a solution is saved.

![Installing ReSharper.AutoFormatOnSave via Visual Studio Extension Manager](http://blog.pedropombeiro.com/wp-content/uploads/2012/10/image_thumb-25255B8-25255D.png)

The profile to use is defined in the box signaled below (ReSharper 6.1 shown here):

![Choosing a default code format profile in ReSharper](http://blog.pedropombeiro.com/wp-content/uploads/2012/10/image_thumb-25255B10-25255D.png)

The addin should be compatible with Visual Studio 2010 or later as well as any version of ReSharper.
