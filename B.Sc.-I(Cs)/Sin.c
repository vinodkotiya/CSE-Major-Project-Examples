	  /*  program to genrat sin series. */

	#include<stdio.h>
	#include<conio.h>
	long int fact(int);
	float power(float,int);
	main()
	{
		float sum,temp,x,pow;
		int  sign,i,n;
		long int factval;
		clrscr();
		printf("\n Enter value x & n :");
		scanf("%f%d",&x,&n);
		i=3;sum=x;sign=1;
	while(i<=n)	
	{
		factval=fact(i);
		pow=power(x,i);
		sign=(-1)*sign;
		temp=sign*pow/factval;
		sum=sum+temp;
		i=i+2;
	}
		printf("sum of x-(x^3)/3!+(x^5)/5!=%f",sum);
		getch();
		return(0);
	}
		long int fact(int m)
	{
		long int value=1,i;
	for(i=1;i<=m;i++)
	{	
		value=value*i;
	}
		return(value);
	}
	float power(float x,int n)
	{
		int j;
		float val=1;
	for(j=1;j<=n;j++)
		{	
		val=val*x;
		return(val);
	}
	}	
**************************************************************************
   /* OUTPUT */
   Enter the value x & n : 2 &5

   Sum of x-(x^3)/3!+(x^5)/5!=1.6333