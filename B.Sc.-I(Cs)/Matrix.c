    /* program to multiply the two matrix */
	#include<stdio.h> 
	#include<conio.h>
	void main()
	{
		int i,j,k,n ;
		int a[3][3],b[3][3],c[3][3];
		clrscr();
		printf("Enter the size of matrix ");
		scanf("%d",&n);
		printf("Enter element in a matrix ");
		for(i=0;i<n;i++)
		{
			for(j=0;j<n;j++)
			{
				scanf("%d",&a[i][j]);
			}
		}
		printf("Enter elements in b matrix");
		for(i=0;i<n;i++)
		{
			for(j=0;j<n;j++)
			{
				scanf("%d",&b[i][j]);
			}
		}
		printf("\nMatrix a is\n");
		for(i=0;i<n;i++)
		{
			for(j=0;j<n;j++)
			{
				printf("\t%d",a[i][j]);
			}
			printf("\n");
		}
		printf("\n matrix b is \n");
		for(i=0;i<n;i++)
		{
			for(j=0;j<n;j++)
			{
				printf("\t%d",b[i][j]);
			}
			printf("\n");
		}
		printf("\n multpal of matrix\n ");
		for(i=0;i<n;i++)

		{
			for(j=0;j<n;j++)
			{
				c[i][j]=0;
				for(k=0;k<n;k++)
				{
					c[i][j]=c[i][j]+a[i][j]*b[j][k];
				}
				printf("\t%d",c[i][j]);
			}
			printf("\n");
			}
			getch();
	}
*************************************************************************************************
	/*OUTPUT*/
	enter the size of matrix 2
	enter element in a matrix 2
	2
	2	
	2	
	enter element in a matrix 2
	2
	2
	2
	matrix a is 
	2	2
	2	2
	matrix b is 
	2	2
	2	2
	multpal of matrix
	8	8	
	8	8


