# Shift_Rotator
 A little excel sheet that that generates rotating work schedules
 
![image](https://user-images.githubusercontent.com/59659182/192485268-dfa69fbb-965e-41ed-a2ae-2bcbefe5e1f5.png)


How to use : 

- ACTIVATE MODIFICATION if asked
- ACTIVATE CONTENT if asked
- Click the Generate button ONLY ONCE, or you will loose previous weeks records and people might be upset if choosen again for a specific shift.
- You can manually finetune the schedule thereafter.  
- You can add a shift slot AT THE END of exisiting shifts (don't forget the shift ID on the first raw, or it wont be taken into account).
- Don't write anything on that first row (shift ID) if you don't intend to add a new shift.
- If you wan't to add a shift slot in between existing ones, to increase number of people in "dinner" shift for instance, or if you want to add a new shift that has more than one people in it, you will have to modify the code by yourself. Sorry. 
- If you have an error message : 
  > the number of people you have is lower than the number of existing tasks. The program doesn't not allow someone to work twice a day, in this case you have to modify the code yourself or do the schedule by hand. 
  > Or you didn't use snake_case "_" in some names.
