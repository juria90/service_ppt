FasdUAS 1.101.10   ��   ��    k             l     ����  r       	  J      
 
     m        �     M K _ s e r v i c e _ t i t l e      m       �    M K _ s e r m o n _ t i t l e      m       �    M K _ d a t e   ��  m       �    n o t   f o u n d��   	 o      ���� 0 
find_texts  ��  ��        l  	  ����  r   	     J   	       !   m   	 
 " " � # # 
���|  ��0 !  $ % $ m   
  & & � ' ' � ��� %  ( ) ( m     * * � + +  2 0 2 0�D   6��   3 0�| )  ,�� , m     - - � . .  x x x��    o      ���� 0 replace_texts  ��  ��     / 0 / l     �� 1 2��   1 , &set find_texts to {"MK_service_title"}    2 � 3 3 L s e t   f i n d _ t e x t s   t o   { " M K _ s e r v i c e _ t i t l e " } 0  4 5 4 l     �� 6 7��   6 $ set replace_texts to {"?? ??"}    7 � 8 8 < s e t   r e p l a c e _ t e x t s   t o   { "���|  ��0 " } 5  9 : 9 l    ;���� ; r     < = < m     > > � ? ?  M a t c h   c a s e = o      ���� 0 match_case_label  ��  ��   :  @ A @ l    B���� B r     C D C m     E E � F F * F i n d   w h o l e   w o r d s   o n l y D o      ���� 0 whole_words_label  ��  ��   A  G H G l    I���� I r     J K J m     L L � M M  R e p l a c e   A l l K o      ���� 0 replaceall_label  ��  ��   H  N O N l    ' P���� P r     ' Q R Q m     # S S � T T  D o n e R o      ���� 0 
done_label  ��  ��   O  U V U l     ��������  ��  ��   V  W X W l     �� Y Z��   Y 5 / delay x : if UI structure changes, wait 1 sec,    Z � [ [ ^   d e l a y   x   :   i f   U I   s t r u c t u r e   c h a n g e s ,   w a i t   1   s e c , X  \ ] \ l     �� ^ _��   ^   otherwise wait 0.5 sec.    _ � ` ` 0   o t h e r w i s e   w a i t   0 . 5   s e c . ]  a b a l     ��������  ��  ��   b  c d c l  (� e���� e O   (� f g f O   .� h i h k   9� j j  k l k r   9 @ m n m m   9 :��
�� boovtrue n 1   : ?��
�� 
pisf l  o p o l  A A��������  ��  ��   p  q r q l  A A�� s t��   s , & click Edit -> Find -> Replace... menu    t � u u L   c l i c k   E d i t   - >   F i n d   - >   R e p l a c e . . .   m e n u r  v w v I  A e�� x��
�� .prcsclicnull��� ��� uiel x n   A a y z y 4   Z a�� {
�� 
menI { m   ] ` | | � } }  R e p l a c e . . . z n   A Z ~  ~ 4   U Z�� �
�� 
menE � m   X Y����   n   A U � � � 4   N U�� �
�� 
menI � m   Q T � � � � �  F i n d � n   A N � � � 4   G N�� �
�� 
menE � m   J M � � � � �  E d i t � 4   A G�� �
�� 
mbar � m   E F���� ��   w  � � � I  f k�� ���
�� .sysodelanull��� ��� nmbr � m   f g���� ��   �  � � � l  l l��������  ��  ��   �  � � � O   l � � � � k   u � � �  � � � r   u ~ � � � 1   u z��
�� 
ects � o      ���� 0 uielems uiElems �  ��� � I   ��� ���
�� .ascrcmnt****      � **** � o    ����� 0 uielems uiElems��  ��   � 4  l r�� �
�� 
cwin � m   p q����  �  � � � l  � ���������  ��  ��   �  � � � r   � � � � � m   � �����   � o      ���� 0 
item_index   �  � � � X   �� ��� � � k   �� � �  � � � l  � ��� � ���   � - ' item_index start from 1 in applescript    � � � � N   i t e m _ i n d e x   s t a r t   f r o m   1   i n   a p p l e s c r i p t �  � � � r   � � � � � [   � � � � � o   � ����� 0 
item_index   � m   � �����  � o      ���� 0 
item_index   �  � � � l  � ���������  ��  ��   �  � � � r   � � � � � n   � � � � � 4   � ��� �
�� 
cobj � o   � ����� 0 
item_index   � o   � ����� 0 replace_texts   � o      ���� 0 replace_text   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   �   Find What:    � � � �    F i n d   W h a t : �  � � � I  � ��� ���
�� .JonspClpnull���     **** � K   � � � � �� ���
�� 
utxt � o   � ����� 0 	find_text  ��  ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � - ' select combo box 1 of window "Replace"    � � � � N   s e l e c t   c o m b o   b o x   1   o f   w i n d o w   " R e p l a c e " �  � � � l  � ��� � ���   � * $ Use copy/paste to handle any chars.    � � � � H   U s e   c o p y / p a s t e   t o   h a n d l e   a n y   c h a r s . �  � � � I  � ��� � �
�� .prcskprsnull���     ctxt � m   � � � � � � �  a v � �� ���
�� 
faal � J   � � � �  ��� � m   � ���
�� eMdsKcmd��  ��   �  � � � l  � ��� � ���   �   keystroke find_text    � � � � (   k e y s t r o k e   f i n d _ t e x t �  � � � I  � ��� ���
�� .sysodelanull��� ��� nmbr � m   � � � � ?�      ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   �   Replace With:    � � � �    R e p l a c e   W i t h : �  � � � I  � ��� ���
�� .JonspClpnull���     **** � K   � � � � �� ���
�� 
utxt � o   � ����� 0 replace_text  ��  ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � - ' select combo box 2 of window "Replace"    � � � � N   s e l e c t   c o m b o   b o x   2   o f   w i n d o w   " R e p l a c e " �  � � � l  � ��� � ���   � S M tell combo box 2 doesn't work, so use tab key to navigate to next combo box.    � � � � �   t e l l   c o m b o   b o x   2   d o e s n ' t   w o r k ,   s o   u s e   t a b   k e y   t o   n a v i g a t e   t o   n e x t   c o m b o   b o x . �  � � � I  � ��� ��
�� .prcskprsnull���     ctxt  l  � ����� I  � �����
�� .sysontocTEXT       shor m   � ����� 	��  ��  ��  ��   �  I  ���
�� .prcskprsnull���     ctxt m   � � �  a v ��	��
�� 
faal	 J   � 

 �� m   � ���
�� eMdsKcmd��  ��    I ���
�� .sysodelanull��� ��� nmbr m   ?�      �    l �~�}�|�~  �}  �|    l �{�{   * $tell combo box 2 of window "Replace"    � H t e l l   c o m b o   b o x   2   o f   w i n d o w   " R e p l a c e "  l �z�z   * $	keystroke "av" using {command down}    � H 	 k e y s t r o k e   " a v "   u s i n g   { c o m m a n d   d o w n }  l �y�y    	# keystroke replace_text    �   2 	 #   k e y s t r o k e   r e p l a c e _ t e x t !"! l �x#$�x  #  	delay 1   $ �%%  	 d e l a y   1" &'& l �w()�w  (  end tell   ) �**  e n d   t e l l' +,+ l �v�u�t�v  �u  �t  , -.- l �s/0�s  / #  move to find what: combo box   0 �11 :   m o v e   t o   f i n d   w h a t :   c o m b o   b o x. 232 I �r4�q
�r .prcskprsnull���     ctxt4 l 5�p�o5 I �n6�m
�n .sysontocTEXT       shor6 m  �l�l 	�m  �p  �o  �q  3 787 l �k�j�i�k  �j  �i  8 9:9 l �h;<�h  ;   log "after replaceall1"   < �== 0   l o g   " a f t e r   r e p l a c e a l l 1 ": >?> I ,�g@�f
�g .prcsclicnull��� ��� uiel@ n  (ABA 4  !(�eC
�e 
butTC o  $'�d�d 0 replaceall_label  B 4  !�cD
�c 
cwinD m   EE �FF  R e p l a c e�f  ? GHG l --�bIJ�b  I   log "after replaceall2"   J �KK 0   l o g   " a f t e r   r e p l a c e a l l 2 "H LML I -2�aN�`
�a .sysodelanull��� ��� nmbrN m  -.�_�_ �`  M OPO l 33�^QR�^  Q   log "after replaceall3"   R �SS 0   l o g   " a f t e r   r e p l a c e a l l 3 "P TUT l 33�]�\�[�]  �\  �[  U VWV l 33�ZXY�Z  X L F get the list of control from below line and find the control you want   Y �ZZ �   g e t   t h e   l i s t   o f   c o n t r o l   f r o m   b e l o w   l i n e   a n d   f i n d   t h e   c o n t r o l   y o u   w a n tW [\[ l 33�Y]^�Y  ]   tell front window   ^ �__ $   t e l l   f r o n t   w i n d o w\ `a` l 33�Xbc�X  b % 	set uiElems to entire contents   c �dd > 	 s e t   u i E l e m s   t o   e n t i r e   c o n t e n t sa efe l 33�Wgh�W  g  	log uiElems   h �ii  	 l o g   u i E l e m sf jkj l 33�Vlm�V  l  	 end tell   m �nn    e n d   t e l lk opo l 33�U�T�S�U  �T  �S  p qrq l 33�Rst�R  s [ U wait for "Powerpoint searched your presentation and made x replacements." dialog box   t �uu �   w a i t   f o r   " P o w e r p o i n t   s e a r c h e d   y o u r   p r e s e n t a t i o n   a n d   m a d e   x   r e p l a c e m e n t s . "   d i a l o g   b o xr vwv l 33�Qxy�Q  x C = Or "We couldn't find what you were looking for." dialog box.   y �zz z   O r   " W e   c o u l d n ' t   f i n d   w h a t   y o u   w e r e   l o o k i n g   f o r . "   d i a l o g   b o x .w {|{ l 33�P}~�P  }   button OK of window 1   ~ � ,   b u t t o n   O K   o f   w i n d o w   1| ��� U  3���� k  <��� ��� l <<�O���O  �   log "repeat"   � ���    l o g   " r e p e a t "� ��� Z  <����N�M� l <J��L�K� =  <J��� n  <F��� 1  BF�J
�J 
pnam� 4  <B�I�
�I 
cwin� m  @A�H�H � m  FI�� ���  �L  �K  � Z  M����G�F� l M^��E�D� I M^�C��B
�C .coredoexnull���     ****� n  MZ��� 4  SZ�A�
�A 
butT� m  VY�� ���  O K� 4  MS�@�
�@ 
cwin� m  QR�?�? �B  �E  �D  � k  a��� ��� r  ao��� n  ak��� 1  gk�>
�> 
pnam� 4  ag�=�
�= 
cwin� m  ef�<�< � o      �;�; 
0 w_name  � ��� I pw�:��9
�: .ascrcmnt****      � ****� o  ps�8�8 
0 w_name  �9  � ��� I x��7��6
�7 .prcsclicnull��� ��� uiel� n  x���� 4  ~��5�
�5 
butT� m  ���� ���  O K� 4  x~�4�
�4 
cwin� m  |}�3�3 �6  � ��� I ���2��1
�2 .sysodelanull��� ��� nmbr� m  ���0�0 �1  � ��/�  S  ���/  �G  �F  �N  �M  � ��.� I ���-��,
�- .sysodelanull��� ��� nmbr� m  ���+�+ �,  �.  � m  69�*�* � ��)� l ���(�'�&�(  �'  �&  �)  �� 0 	find_text   � o   � ��%�% 0 
find_texts   � ��� l ���$�#�"�$  �#  �"  � ��� I ���!�� 
�! .prcsclicnull��� ��� uiel� n  ����� 4  ����
� 
butT� o  ���� 0 
done_label  � 4  ����
� 
cwin� m  ���� ���  R e p l a c e�   � ��� I �����
� .sysodelanull��� ��� nmbr� m  ���� �  � ��� l ������  �  �  � ��� I �����
� .ascrcmnt****      � ****� m  ���� ���  F i n a l l y   d o n e .�  �   i 4   . 6��
� 
prcs� m   2 5�� ��� ( M i c r o s o f t   P o w e r P o i n t g m   ( +���                                                                                  sevs  alis    \  Macintosh HD                   BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ��  ��   d ��� l     ����  �  �  �       ����  � �
� .aevtoappnull  �   � ****� ����
���	
� .aevtoappnull  �   � ****� k    ���  ��  ��  9��  @��  G��  N��  c��  �  �
  � �� 0 	find_text  � =    �� " & * -� >� E� L� S� ���������� ��� � |�������������������������� ������� �����E���������������� � 0 
find_texts  � 0 replace_texts  � 0 match_case_label  � 0 whole_words_label  � 0 replaceall_label  �  0 
done_label  
�� 
prcs
�� 
pisf
�� 
mbar
�� 
menE
�� 
menI
�� .prcsclicnull��� ��� uiel
�� .sysodelanull��� ��� nmbr
�� 
cwin
�� 
ects�� 0 uielems uiElems
�� .ascrcmnt****      � ****�� 0 
item_index  
�� 
kocl
�� 
cobj
�� .corecnte****       ****�� 0 replace_text  
�� 
utxt
�� .JonspClpnull���     ****
�� 
faal
�� eMdsKcmd
�� .prcskprsnull���     ctxt�� 	
�� .sysontocTEXT       shor
�� 
butT�� 
�� 
pnam
�� .coredoexnull���     ****�� 
0 w_name  �	������vE�O�����vE�O�E�O�E�O�E` Oa E` Oa �*a a /�e*a ,FO*a k/a a /a a /a k/a a /j Okj O*a k/ *a  ,E` !O_ !j "UOjE` #O�[a $a %l &kh  _ #kE` #O�a %_ #/E` 'Oa (�lj )Oa *a +a ,kvl -Oa .j Oa (_ 'lj )Oa /j 0j -Oa 1a +a ,kvl -Oa .j Oa /j 0j -O*a a 2/a 3_ /j Okj O pa 4kh*a k/a 5,a 6  M*a k/a 3a 7/j 8 5*a k/a 5,E` 9O_ 9j "O*a k/a 3a :/j Okj OY hY hOkj [OY��OP[OY��O*a a ;/a 3_ /j Okj Oa <j "UU ascr  ��ޭ