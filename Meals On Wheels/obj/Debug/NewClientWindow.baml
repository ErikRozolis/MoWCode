   M S B A M L   `   `   `  ����J  FMeals On Wheels, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Meals_On_WheelsMeals_On_Wheels  5   	   %    Meals_On_Wheels.NewClientWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:Meals_On_Wheels     5   	     � Title$  NewClientWindow��5   	   $	��500��6!   $	��800��6.   .��5	      � .+�5
   
   �� -   5         ��Name$ ClientNameTextBox��5
      $
��Left=�6-   $��40��6H   $��	250,0,0,0q�6T    �� TextWrapping$
 Wrap=�6g   $	��Top=�6{   $	��500��6�   $��20(�6�   �� 5   
   -   5        $ ClientAddressTextBox��5      $
��Left=�60   $��40��6K   $��
250,50,0,0q�6W   $
 Wrap=�6k   $	��Top=�6   $	��500��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��	0,0,250,0q�6/   $
��Wrap=�6B    �� TextAlignment$ Right=�6V   $��Client Name:��6l   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��
0,50,250,0q�6/   $
��Wrap=�6C   $ Right=�6W   $��Client Address:��6m   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��	0,100,0,0q�6/   $
��Wrap=�6B   $ Right=�6V   $��# of Hot Meals:��6l   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��	0,150,0,0q�6/   $
��Wrap=�6B   $ Right=�6V   $��# of Cold Meals:��6l   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��	0,200,0,0q�6/   $
��Wrap=�6B   $ Right=�6V   $��Special Instructions:��6l   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   -   5        $ HotMealsTextBox��5      $
��Left=�6+   $��40��6F   $��250,100,0,0q�6R   $
 Wrap=�6g   $	��Top=�6{   $	��500��6�   $��20(�6�   �� 5   
   -   5        $ ColdMealsTextBox��5      $
��Left=�6,   $��40��6G   $��250,150,0,0q�6S   $
 Wrap=�6h   $	��Top=�6|   $	��500��6�   $��20(�6�   �� 5   
   -   5        $  SpecialInstructionsTextBox��5      $
��Left=�66   $��40��6Q   $��250,200,0,0q�6]   $
 Wrap=�6r   $	��Top=�6�   $	��500��6�   $��20(�6�   �� 5   
   $
��Left=�6   $��	0,250,0,0q�6/   $
��Wrap=�6B   $ Right=�6V   $��Phone Number:��6l   $	��Top=�6�   $	��200��6�   $��40��6�   $��20(�6�   �� 5   
   -   5        $ PhoneNumberTextBox��5      $
��Left=�6.   $��40��6I   $��250,250,0,0q�6U   $
 Wrap=�6j   $	��Top=�6~   $	��500��6�   $��20(�6�   �� 5   
   -   5         ��Name$ CreateClient��5      $��Create/Edit Client��6'   $��Center=�6D   $��0,0,0,30q�6a   $��Bottom=�6s   $	��200��6�   $��20(�6�   $��40��6�   6�   5      5      