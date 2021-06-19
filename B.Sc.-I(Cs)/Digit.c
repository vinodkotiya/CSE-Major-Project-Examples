		/*program to print given number in revers order*/

		#include<stdio.h>
		#include<conio.h>
		main()
		{
			int number,digit;
			clrscr();
			printf("\n enter number");
			scanf("%d",&number);
			while(number>0)
		{
			digit=number%10;
			number=number/10;
			printf("%d",digit);
		}
		getch();
		return 0;
		}

**************************************************************************

		/*OUTPUT*/
		Enter number :348
		843

