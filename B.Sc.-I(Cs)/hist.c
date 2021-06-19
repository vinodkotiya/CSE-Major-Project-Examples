		/*PROGRAM FOR GENERATING HISTOGRAM*/
	#include<stdio.h>
	#include<conio.h>
	main()
	{
		int i,j,n;
		clrscr();
		printf("\nHow many rows you want to print:");
		scanf("%d",&n);
		for(i=1; i<=n; i++)
		{
		for(j=1; j<=i; j++)
		{
		printf("*\t");
		}
		printf("\n");
		}
		getch();
	}


*************************************************************
	/*OUTPUT*/
	How many rows you want to print 4
	*
	*  *
	*  *  *
	*  *  *  *
	How many rows you want to print  2
	*
	* *
