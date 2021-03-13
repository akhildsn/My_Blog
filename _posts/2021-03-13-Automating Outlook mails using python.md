---
layout: category-post
title:  "Automating Outlook e-mails using Python (O365)"
date:   2021-2-17
categories: writing
---
###Context

I was recently tasked with reaching out to people who demoed products on our website and see if there are any genuine leads. The only infomation I had was their email, what product they treid and the country they were from. I realised I have to first reach out to all of them thorugh mail and take it forward when I get a response. With limited information available, I had to use a  generic template for the mail content and send a bunch of emails. I wasn't eaxctly thrilled about the idea and since I've been learning python recently, I thought why not automate this? Now, my interaction with automation before this point was only with my old laptop that used to shut down on it's own. So, this is quite new to me and you will understand that one you go through the tutotial. In the spirit of embracing the noobness, I may explain some very basic stuff I figured out as well, this is also to encourage anyone who has no experience in python to try this because <b>If I can do it, anyone can do it</b>. 

#Tutorial
_ _
In this post, we will be looking at setting up a script to send out emails from Outlook using O365 library in python. There are alternate ways to do this, especially for your  personal account using SMTP (Simple Mail Transfer Protocol) which is a built-in library in python. You can refer to this tutorial to set it up. But if you want to set it up on your work account (in my case Outlook), that won't work as there are additional layers of security. Which is why, we will be using O365 library for it. I will only get into setting up basic capabiilties like sending emails but there's a lot of stuff you can do using O365. You'll find more information on that<link rel="stylesheet" type="text/css" href="https://github.com/O365/python-o365"> here. </link>  

##_Setting up_

We first need to install O365 library to do so just run ```pip install O365``` wherever you are running python (I'm assuming you've already. If not, I suggest using jupyter notebook - Anaconda). 




Hi Everyone! In this blog post, I am going to explain how to setup automated mailing script using python for your Outlook work account. 