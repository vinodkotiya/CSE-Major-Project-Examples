	/*PROGRAM TO CHECK WEATHER A GIVEN WORD IS PALINDROME */

	#include<stdio.h>
	#include<conio.h>
	#include<string.h>
	main()
	{
		char n[30],c[30];
		int p;
		clrscr();
		printf("\n\n\t enter the string\n");
		scanf("%s",n);
		strcpy(c,n);
		strrev(n);
		p=strcmp(c,n);
		if(p==0)
		{
			printf("\n\n\tstring is a palindrome\n");
		}
		else
		{
			printf("\n\n\tstring is not palindrome\n");
		}
		getch();
		return 0;
	}


***************************************************************************
		/*OUTPUT*/
		ENTER THE STRING
		MADAM
		STRING IS A PALINDROME
