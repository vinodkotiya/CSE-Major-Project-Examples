	/* WRITE A PROGRAM TO CALCULAT THE FACTONAL OF A GIVEN NUMBER*/
	
		#include <stdio.h>
		#include <conio.h>
		void main ()
		{
			  int number,fact;
			 clrscr();
			 printf("\n enter number");
			 scanf("%d",&number);
			 fact=factorial(number);
			 printf("\n the factorial of %d is %d",number,fact);
			 printf("\n \n ");
			 getch();
			 }
			 factorial(number)
			 int number;
			 {
			 int a;
			 if(number==0)
			 return (1);
			else
			a=number*factorial(number-1);
			return(a);
		}
			







	/*OUTPUT*/
	Enter number:3
	The factorial of 3 is 6

	











