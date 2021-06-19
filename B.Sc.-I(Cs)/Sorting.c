	/*program for sorting an ordered array*/
	
	#include<stdio.h>
	#include<conio.h>
	main()
	{
		int a[4],j=0,n,i,k=1,t;
		clrscr();
		printf("\n\t how many number you want to input in an array :");
		scanf("%d",&n);
		for(i=0;i<n;i++)
		{
			printf("Enter the %d number in array y :",i+1);
			scanf("%d",&a[i]);
		}
			while(k==1 &&j<n)
		{
			k=0;
			for(i=0;i<(n-1);i++)
		{
			if(a[i]>a[i+1])
		{
			k=1;
			t=a[i];
			a[i]=a[i+1];
			a[i+1]=t;
			}
			}
			j++;	
			}
			printf("\t\n the sorted order is \n\n ");
		for(i=0;i<n;i++)
		{
		printf("\n\t a[%d]=%d",i+1,a[i]);
		}
		getch();
		return 0;
	}

*************************************************************************
	/*OUTPUT*/
	HOW MANY NUMBER YOU WANT TO INPUT IN ARRAY :4
	Enter the 1 number in array Y :3
	Enter the 2 number in array Y :2
	Enter the 3 number in array Y :1
	Enter the 4 number in array Y :5
		
	THE STORED ORDER IS
	a[1]=3
	a[2]=2
	a[3]=1
	a[4]=5
