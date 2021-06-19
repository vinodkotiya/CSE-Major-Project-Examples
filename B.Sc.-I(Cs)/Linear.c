
		/* PROGRAM FOR LINEAR SEARCH.*/

	#include<stdio.h>
	#include<conio.h>
	void main()
	{
		int num[5],i ,tnum,flag;
		clrscr();
		printf("\n Enter elument in array :");
		for(i=0;i<5;i++)
		{
		scanf("%d",& num[i]);
		}
		flag=0;
		printf("Enter num to be searched  :");
		scanf("%d",&tnum);
		for(i=0;i<5;i++)
		{
		if(num[i]==tnum)
		flag=1;
		}
		if(flag==1)
		printf("Number exist");
		else
		printf("Number does not exist");
		getch();	
	}
	
**************************************************************************
  	/* OUTPUT */
	   Enter elument in array :7
	   8
	   7
	   3
	   Enter number to be search : 3
   		Number exist.
	
