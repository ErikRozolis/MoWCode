   M S B A M L   `   `   `  ����J  FMeals On Wheels, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null1clr-namespace:Meals_On_WheelsMeals_On_Wheels  5   	        Meals_On_Wheels.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   /localclr-namespace:Meals_On_Wheels     5   	     � Title$  
MainWindow��5   	   $	��600��6   $
��1000��6)   .��5	      � .+�5
   
   � -   5         �Name$ DisplayGrid��5
      $��0q�6$   ���	#FFF3F3F3.+�5      �� -   5         ��Name$ RouteSelectionBox��5      $��	30,30,0,0q�62   $	��200��6E   $	��Top=�6Q   $
��Left=�6i   $��Center=�6�   $��40��6�   6�   �� 5      $��30,100q�6   $��Center=�6+   $��Center=�6O   .��5      � -   5        $ 
ClientList��5      ���	#FFBDE7FF � ShowGridLines . $	��900��6W   $��Center=�6c    � RowDefinitions 5      �� 5      $��40 �6(   5      5      �� 5      -   5         ��Name$ NewRouteButton��5      $��Create New Route��6-   $	��Top=�6H   $	��200��6`   $��	0,30,30,0q�6l   $��40��6   $��Right=�6�   $��Center=�6�   $��20(�6�   6�   �� 5      -   5        $ DeleteRouteButton��5      $��X��60   $	��Top=�6<   $��40��6T   $��
230,30,0,0q�6_   $��40��6s   $��Center=�6   $��24(�6�   $
��Left=�6�   
���  ��$��	ExtraBold"�6�   �� 5      -   5        $ GoogleMapsButton��5      $��Center=�6/   $��0,0,0,30q�6L   $��Bottom=�6^   $	��200��6y   $��40��6�   $��Center=�6�   $��20(�6�   $��Generate Excel��6�   �� 5      -   5         ��Name$ LoadingText��5      $��Center=�6-   $��0,30,0,0q�6J   $
��Wrap=�6\   $��
Loading...��6p   $	��Top=�6�   $��36(�6�   ���	#FFFF0707$
��Bold"�6�   5      5      5      