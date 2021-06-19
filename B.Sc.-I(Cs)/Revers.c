	    /* Program for reversing an array.*/
	
	#include <stdio.h>
	#include <conio.h>
	main()
	{
		  int i,a[5];
		  clrscr();
		  printf("Enter array value  :\n");
		  for(i=0;i<5;i++)
		  {
		  scanf("%d",&a[i]);
		  }
		  printf("\nArray value are  :  \n");
		  for(i=4;i>=0;i--)
		  {
		   printf("%d\n",a[i]);
		  }
		  getch();
	}

****************************************************************************
 		 /* OUTPUT */
		 Enter array value :
		 1
		 2
		 3
		 4
		 5
		 Array values are :
		 5
		 4
		 3
		 2
		 1
			
