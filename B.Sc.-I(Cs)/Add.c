
      /* PROGRAM TO ADD DIGIT OF A NUMBER */
	
	#include<stdio.h>
	#include<conio.h>
	int sum_of_digit(int n)
	{
		int sum,rem;
		sum=0;
		while(n!=0)
		{
			rem=n%10;
			n=n/10;
			sum=sum+rem;
		}
		return(sum);
	}


	void main()
	{
		int n,sum,ans;
		int sum_of_digit();
		clrscr();
		printf("Enter number:");
		scanf("%d",&n);
		ans=sum of digit(n);
		printf("Sum of digit is %d\n:",ans);
		getch();
	}



*******************************************************************

	/* OUTPUT */
	Enter number :45
	Sum of digit:9
 



