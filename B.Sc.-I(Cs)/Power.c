

	#include<stdio.h>
	#include<conio.h>
	#include<math.h>
	main()
	{
		int x,n,p,i;
		p=1;
		clrscr();
		printf("\n enter the value of x :");
		scanf("%d",&x);
		printf("\n enter the value of n :");
		scanf("%d",&n);
		for(i=1;i<=n;i++)
		{
		p=p*x;
		}
		printf("\n\n%dto the power %d is %d",x,n,p);
		getch();
		return 0;
	}

************************************************************

	/* output */
	enter the value of x : 3
	enter the value of n : 3
	3 to the power 3 is 27



