This simple tool can be used to launch a set of web clients that perform a GET operation on the specified URL for the specified number of times, in a different set of threads. 

I created this to be able to see the AWS ELB in action.

I had configured 2 EC2 instances and added them behind an Application Load balancer. One was an Ubuntu based EC2 and the other one an Amazon linux based one. Both were in Virginia. By launching the application, it spawned off a few threads that iterated through a loop and sent a set of consecutive requests to the ALB. In the output stored in the Results folder, we can see the change happening between the EC2 instances. 

TODO:
1. A UI front end of the tool
2. If a UI is an overkill, to move the parameters into an app.config file