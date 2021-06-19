	/* Program for finding GCD.*/

	#include <stdio.h>
	#include <conio.h>
	main()
	{
		int a,b,k;
		clrscr();
		printf("\n Enter the first number:");
		scanf("%d",&a);
		printf("\n Enter second number :");
		scanf("%d",&b);
		while(b!=0)
			 {
				k=a%b;
				a=b;
				b=k;
			 }

		printf("\n Gcd=%d",a);
		getch();
	 }




*****************************************************************************
	/* OUTPUT */
	Enter first number:25
	Enter second number:20
	GCD =5
