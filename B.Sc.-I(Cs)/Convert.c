

		


           /*PROGRAM TO CONVERT BINARY INTO DECIMAL NUMBER */	
	
		#include<stdio.h>
		#include<conio.h>
		#include<math.h>
		main()
		{
			int i=0,k=0,r,n,p,m=0;
			clrscr();
			printf("\n\ninput a binary number \n\n");
			scanf("%d",&n);
			p=n;
			while(n!=0)
			{
			r=n%10;
			k=r*pow(2,i);
			m+=k;
			n=n/10;
			i++;
			}
			printf("\n conversion of \n");
			printf("\n binary number =%d\t decimal number =%d\n",p,m);
			getch();
			return(0);
		}



****************************************************************************
		/*OUTPUT*/
		INPUT A BINARY NUMBER:
			101
		COVERSION OF
		BINARy NUMBER=101       DECIMAL NUMBER=5


