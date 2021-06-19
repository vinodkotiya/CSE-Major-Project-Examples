	/* Program to check weather a number is prime or not*/

	#include <stdio.h>							
	#include <conio.h>
	main()
	{
		int num,i,count=1;
		clrscr();
		printf("\nEntar a number :");
		scanf("%d",&num);
		if(num==2)
		{
		printf("\n Number is prime");
		getch();
		exit(0);
		}			
		for(i=2;i<num;i++)
		{
		if((num%i)==0)
		{
		count=0;
		break;
		}
		}
		if(count!=0)
		{
		printf("\nNumber is prime");
		}
		else
		{
		printf("\n Number is not prime");
				}
		getch();
		return 0;
	}	
****************************************************************
	/* OUTPUT */
	Enter a number :5
	Number is prime
	Enter a number :6
	Number is not prime
