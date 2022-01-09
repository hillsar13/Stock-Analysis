# Stock-Analysis

## Overview of Project
The purpose of this project was to determine whether refactoring a previously completed code would make the VBA script run faster. The client in question is interested in further analysis of the entire stock market as oppose to a few individual stocks. The goal was  to see if we could create a more efficient script to provide this possibility for the client. 

## Results: 
  In support of our original hypothesis, the refactored code allowed the script to run faster for both 2017 & 2018 stocks. 
  
  The performance results from the original code for 2017 & 2018:
  
  <img width="258" alt="Original_VBA_Challenge_2017" src="https://user-images.githubusercontent.com/95551195/148663549-3d7914e2-1687-48fd-8b98-0eb001bc616b.png">  <img width="260" alt="Original_VBA_Challenge_2018" src="https://user-images.githubusercontent.com/95551195/148663583-5c843a79-d979-4e0c-9896-f7feefc581e5.png">

   The performance results from the refactored code for 2017 & 2018:
    
<img width="258" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/95551195/148663550-26605a2b-9032-4055-9025-d55c38773346.png">  <img width="257" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/95551195/148663587-f84bf187-f24a-4499-962f-4d60a673abfa.png">

  Our analysis was able to cut the performance time of the script in more than half the original time. By adjusting the code to omit the need for a loop within a     loop, we were able to achieve these results.
  
  An example of code from the original script:
  
  <img width="530" alt="Original VBA Script" src="https://user-images.githubusercontent.com/95551195/148663818-cd464c1d-ba99-4d45-a084-62d82bb47ad7.png">

  
  An example of code from the refactored script:
  
  <img width="651" alt="Refactored VBA Script" src="https://user-images.githubusercontent.com/95551195/148663819-9cc2b212-8216-4ae0-b7d4-504bc3634773.png">

  
  
## Summary: 

There are both advantages and disadvantages to refactoring code. An advantage as reflected in this analysis is the opportunity to capitzlie on performance time and efficacy for a script. In the circumstance with our client, this also may allow him to get the information they are seeking at a much faster rate. In essence, we were able to improve the internal structure of the code without altering the output.  A disadvantage of this process, ironically, can be the time needed to refactor a code. For this client, the time saving nature of this script are beneficial, but there may be less time or demand sensitive results needed in a code, rendering the refactor unnecessary. It also may better serve the user of the VBA to start fresh with code as oppose to attempt to revamp it.

## How do these pros and cons apply to factoring the original VBA script?

These pros and cons apply directly to factoring the original VBA script. Upon successful refactoring, the analysis above shows that the script can be altered to improve performance time. However, as a novice to VBA, the amount of time required to refactor far outweighed creating the original script. However, this user error will likely decrease with time as the user becomes more well versed in VBA. 
