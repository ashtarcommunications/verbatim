FasdUAS 1.101.10   ��   ��    k             l     ��  ��    $  Verbatim Scripts for Mac Word     � 	 	 <   V e r b a t i m   S c r i p t s   f o r   M a c   W o r d   
  
 l     ��  ��    #  Copyright � 2023 Aaron Hardy     �   :   C o p y r i g h t   �   2 0 2 3   A a r o n   H a r d y      l     ��  ��    "  https://paperlessdebate.com     �   8   h t t p s : / / p a p e r l e s s d e b a t e . c o m      l     ��  ��    "  support@paperlessdebate.com     �   8   s u p p o r t @ p a p e r l e s s d e b a t e . c o m      x     �� ����    4    �� 
�� 
frmk  m       �    F o u n d a t i o n��         x    �� !����   ! 2   ��
�� 
osax��      " # " l     ��������  ��  ��   #  $ % $ i     & ' & I      �� (����  0 runshellscript RunShellScript (  )�� ) o      ���� 0 scripttorun ScriptToRun��  ��   ' I    �� *��
�� .sysoexecTEXT���     TEXT * o     ���� 0 scripttorun ScriptToRun��   %  + , + l     ��������  ��  ��   ,  - . - i     / 0 / I      �� 1���� 0 
openfolder 
OpenFolder 1  2�� 2 o      ���� 0 
folderpath 
FolderPath��  ��   0 I    	�� 3��
�� .sysoexecTEXT���     TEXT 3 b      4 5 4 b      6 7 6 m      8 8 � 9 9  o p e n   ' 7 o    ���� 0 
folderpath 
FolderPath 5 m     : : � ; ;  '��   .  < = < l     ��������  ��  ��   =  > ? > i    " @ A @ I      �������� *0 getfolderfromdialog GetFolderFromDialog��  ��   A L      B B c      C D C n      E F E 1    ��
�� 
psxp F l     G���� G I    ���� H
�� .sysostflalis    ��� null��   H �� I J
�� 
prmp I m     K K � L L " S e l e c t   t h e   f o l d e r J �� M��
�� 
dflc M l    N���� N c     O P O n     Q R Q 1   	 ��
�� 
psxp R l   	 S���� S I   	�� T��
�� .earsffdralis        afdr T m    ��
�� afdrcusr��  ��  ��   P m    ��
�� 
TEXT��  ��  ��  ��  ��   D m    ��
�� 
TEXT ?  U V U l     ��������  ��  ��   V  W X W i   # & Y Z Y I      �������� &0 getfilefromdialog GetFileFromDialog��  ��   Z L      [ [ c      \ ] \ n      ^ _ ^ 1    ��
�� 
psxp _ l     `���� ` I    ���� a
�� .sysostdfalis    ��� null��   a �� b c
�� 
prmp b m     d d � e e  S e l e c t   t h e   f i l e c �� f��
�� 
dflc f l    g���� g c     h i h n     j k j 1   	 ��
�� 
psxp k l   	 l���� l I   	�� m��
�� .earsffdralis        afdr m m    ��
�� afdrcusr��  ��  ��   i m    ��
�� 
TEXT��  ��  ��  ��  ��   ] m    ��
�� 
TEXT X  n o n l     ��������  ��  ��   o  p q p i   ' * r s r I      �� t���� .0 getsubfoldersinfolder GetSubFoldersInFolder t  u�� u o      ���� 0 
folderpath 
FolderPath��  ��   s O     3 v w v k    2 x x  y z y r     { | { m     } } � ~ ~   | o      ���� 0 r   z   �  r     � � � n     � � � 2   ��
�� 
cfol � 4    �� �
�� 
cfol � o   
 ���� 0 
folderpath 
FolderPath � o      ���� 0 	myfolders 	myFolders �  � � � X    / ��� � � r   ! * � � � b   ! ( � � � l  ! & ����� � c   ! & � � � b   ! $ � � � o   ! "���� 0 r   � o   " #���� 0 f   � m   $ %��
�� 
TEXT��  ��   � m   & ' � � � � �  \ n � o      ���� 0 r  �� 0 f   � o    ���� 0 	myfolders 	myFolders �  � � � l  0 0��������  ��  ��   �  ��� � L   0 2 � � o   0 1���� 0 r  ��   w m      � ��                                                                                  MACS  alis    :  	HardyBook                      BD ����
Finder.app                                                     ����            ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p   	 H a r d y B o o k  &System/Library/CoreServices/Finder.app  / ��   q  � � � l     ��������  ��  ��   �  � � � i   + . � � � I      �� ����� $0 getfilesinfolder GetFilesInFolder �  ��� � o      ���� 0 	posixpath 	POSIXPath��  ��   � I    	�� ���
�� .sysoexecTEXT���     TEXT � b      � � � b      � � � m      � � � � �  f i n d   - E   � o    ���� 0 	posixpath 	POSIXPath � m     � � � � � |   - i r e g e x   ' . * / [ ^ ~ ] [ ^ / ] * \ . ( d o c x | d o c | d o c m | d o t | d o t m ) $ '   - m a x d e p t h   1��   �  � � � l     ��������  ��  ��   �  � � � i   / 2 � � � I      �� ����� 0 killfileonmac KillFileOnMac �  ��� � o      ���� 0 filename FileName��  ��   � O      � � � I   �� ���
�� .sysoexecTEXT���     TEXT � b     � � � m     � � � � �  r m   � n    
 � � � 1    
��
�� 
strq � n     � � � 1    ��
�� 
psxp � o    ���� 0 filename FileName��   � m      � ��                                                                                  MACS  alis    :  	HardyBook                      BD ����
Finder.app                                                     ����            ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p   	 H a r d y B o o k  &System/Library/CoreServices/Finder.app  / ��   �  � � � l     �������  ��  �   �  � � � i   3 6 � � � I      �~ ��}�~ "0 killfolderonmac KillFolderOnMac �  ��| � o      �{�{ 0 
foldername 
FolderName�|  �}   � O      � � � I   �z ��y
�z .sysoexecTEXT���     TEXT � b     � � � m     � � � � �  r m   - f r   � n    
 � � � 1    
�x
�x 
strq � n     � � � 1    �w
�w 
psxp � o    �v�v 0 
foldername 
FolderName�y   � m      � ��                                                                                  MACS  alis    :  	HardyBook                      BD ����
Finder.app                                                     ����            ����  
 cu             CoreServices  )/:System:Library:CoreServices:Finder.app/    
 F i n d e r . a p p   	 H a r d y B o o k  &System/Library/CoreServices/Finder.app  / ��   �  � � � l     �u�t�s�u  �t  �s   �  � � � i   7 : � � � I      �r ��q�r 0 startrecord StartRecord �  ��p � o      �o�o 0 paramstring paramString�p  �q   � k      � �  � � � O      � � � k     � �  � � � I   	�n�m�l
�n .MVWRnwarnull��� ��� null�m  �l   �  ��k � I  
 �j ��i
�j .MVWRstarnull���     docu � 4   
 �h �
�h 
docu � m     � � � � �  A u d i o   R e c o r d i n g�i  �k   � m      � ��                                                                                  mgvr  alis    R  	HardyBook                      BD ����QuickTime Player.app                                           ����            ����  
 cu             Applications  +/:System:Applications:QuickTime Player.app/   *  Q u i c k T i m e   P l a y e r . a p p   	 H a r d y B o o k  (System/Applications/QuickTime Player.app  / ��   �  � � � l   �g�f�e�g  �f  �e   �  ��d � L     � � m    �c�c �d   �  � � � l     �b�a�`�b  �a  �`   �  � � � i   ; > � � � I      �_ ��^�_ 0 
saverecord 
SaveRecord �  ��] � o      �\�\ 0 filename FileName�]  �^   � k     F � �  � � � r      � � � 4     �[ �
�[ 
psxf � o    �Z�Z 0 filename FileName � o      �Y�Y 0 
exportpath 
exportPath �  � � � O    C � � � k    B � �    I   �X�W
�X .MVWRstopnull���     docu 4    �V
�V 
docu m     �  A u d i o   R e c o r d i n g�W    r    %	 n    #

 4    #�U
�U 
cobj m   ! "�T�T�� l    �S�R 6     2   �Q
�Q 
docu E     1    �P
�P 
pnam m     �  U n t i t l e d�S  �R  	 o      �O�O 0 doc    I  & 2�N
�N .MVWRexponull���     docu o   & '�M�M 0 doc   �L
�L 
kfil 4   ( ,�K
�K 
file o   * +�J�J 0 
exportpath 
exportPath �I�H
�I 
expp m   - . �  A u d i o   O n l y�H    I  3 <�G !
�G .coreclosnull���     obj   o   3 4�F�F 0 doc  ! �E"�D
�E 
savo" m   5 8�C
�C savono  �D   #�B# I  = B�A�@�?
�A .aevtquitnull��� ��� null�@  �?  �B   � m    $$�                                                                                  mgvr  alis    R  	HardyBook                      BD ����QuickTime Player.app                                           ����            ����  
 cu             Applications  +/:System:Applications:QuickTime Player.app/   *  Q u i c k T i m e   P l a y e r . a p p   	 H a r d y B o o k  (System/Applications/QuickTime Player.app  / ��   � %&% l  D D�>�=�<�>  �=  �<  & '�;' L   D F(( m   D E�:�: �;   � )*) l     �9�8�7�9  �8  �7  * +,+ i   ? B-.- I      �6�5�4�6 20 gethorizontalwindowsize GetHorizontalWindowSize�5  �4  . k     ,// 010 r     232 l    4�3�24 n    565 I    �1�0�/�1 0 visibleframe visibleFrame�0  �/  6 n    787 I    �.�-�,�. 0 
mainscreen 
mainScreen�-  �,  8 n    9:9 o    �+�+ 0 nsscreen NSScreen: m     �*
�* misccura�3  �2  3 o      �)�) *0 visiblescreenbounds visibleScreenBounds1 ;<; r    =>= n    ?@? 4    �(A
�( 
cobjA m    �'�' @ n    BCB 4    �&D
�& 
cobjD m    �%�% C o    �$�$ *0 visiblescreenbounds visibleScreenBounds> o      �#�# 0 leftpos leftPos< EFE r    !GHG n    IJI 4    �"K
�" 
cobjK m    �!�! J n    LML 4    � N
�  
cobjN m    �� M o    �� *0 visiblescreenbounds visibleScreenBoundsH o      �� 	0 width  F OPO l  " "����  �  �  P QRQ l  " "�ST�  S F @ AppleScriptTask wants a string so return values comma delimited   T �UU �   A p p l e S c r i p t T a s k   w a n t s   a   s t r i n g   s o   r e t u r n   v a l u e s   c o m m a   d e l i m i t e dR V�V L   " ,WW c   " +XYX b   " )Z[Z b   " '\]\ l  " %^��^ c   " %_`_ o   " #�� 	0 width  ` m   # $�
� 
TEXT�  �  ] m   % &aa �bb  ,[ o   ' (�� 0 leftpos leftPosY m   ) *�
� 
TEXT�  , cdc l     ����  �  �  d efe i   C Fghg I      �i�� 0 splitstring SplitStringi jkj o      �� 0 thebigstring TheBigStringk l�l o      �
�
  0 fieldseparator fieldSeparator�  �  h k     mm non O     pqp k    rr sts r    	uvu 1    �	
�	 
txdlv o      �� 0 oldtid oldTIDt wxw r   
 yzy o   
 ��  0 fieldseparator fieldSeparatorz 1    �
� 
txdlx {|{ r    }~} n    � 2   �
� 
citm� o    �� 0 thebigstring TheBigString~ o      �� 0 theitems theItems| ��� r    ��� o    �� 0 oldtid oldTID� 1    � 
�  
txdl�  q 1     ��
�� 
ascro ���� L    �� o    ���� 0 theitems theItems��  f ��� l     ��������  ��  ��  � ��� i   G J��� I      ������� 0 activatetimer ActivateTimer� ���� o      ���� 0 timerapp TimerApp��  ��  � O    ��� I   ������
�� .miscactvnull��� ��� null��  ��  � 4     ���
�� 
capp� o    ���� 0 timerapp TimerApp� ��� l     ��������  ��  ��  � ��� i   K N��� I      ������� (0 getfromcitecreator GetFromCiteCreator� ���� o      ���� 0 paramstring paramString��  ��  � k     1�� ��� O     &��� k    %�� ��� I   	������
�� .miscactvnull��� ��� null��  ��  � ��� I  
 �����
�� .sysodelanull��� ��� nmbr� m   
 ���� ��  � ��� O   ��� I   ����
�� .prcskprsnull���     ctxt� m    �� ���  c� �����
�� 
faal� J    �� ��� m    ��
�� eMdsKctl� ���� m    ��
�� eMdsKopt��  ��  � m    ���                                                                                  sevs  alis    V  	HardyBook                      BD ����System Events.app                                              ����            ����  
 cu             CoreServices  0/:System:Library:CoreServices:System Events.app/  $  S y s t e m   E v e n t s . a p p   	 H a r d y B o o k  -System/Library/CoreServices/System Events.app   / ��  � ���� I    %�����
�� .sysodelanull��� ��� nmbr� m     !���� ��  ��  � m     ���                                                                                  rimZ  alis    8  	HardyBook                      BD ����Google Chrome.app                                              ����            ����  
 cu             Applications  !/:Applications:Google Chrome.app/   $  G o o g l e   C h r o m e . a p p   	 H a r d y B o o k  Applications/Google Chrome.app  / ��  � ���� O   ' 1��� I  + 0������
�� .miscactvnull��� ��� null��  ��  � m   ' (���                                                                                  MSWD  alis    <  	HardyBook                      BD ����Microsoft Word.app                                             ����            ����  
 cu             Applications  "/:Applications:Microsoft Word.app/  &  M i c r o s o f t   W o r d . a p p   	 H a r d y B o o k  Applications/Microsoft Word.app   / ��  ��  � ���� l     ��������  ��  ��  ��       ��������������������  � ������������������������������
�� 
pimr��  0 runshellscript RunShellScript�� 0 
openfolder 
OpenFolder�� *0 getfolderfromdialog GetFolderFromDialog�� &0 getfilefromdialog GetFileFromDialog�� .0 getsubfoldersinfolder GetSubFoldersInFolder�� $0 getfilesinfolder GetFilesInFolder�� 0 killfileonmac KillFileOnMac�� "0 killfolderonmac KillFolderOnMac�� 0 startrecord StartRecord�� 0 
saverecord 
SaveRecord�� 20 gethorizontalwindowsize GetHorizontalWindowSize�� 0 splitstring SplitString�� 0 activatetimer ActivateTimer�� (0 getfromcitecreator GetFromCiteCreator� ����� �  ��� �����
�� 
cobj� ��   �� 
�� 
frmk��  � �����
�� 
cobj� ��   ��
�� 
osax��  � �� '����������  0 runshellscript RunShellScript�� ����� �  ���� 0 scripttorun ScriptToRun��  � ���� 0 scripttorun ScriptToRun� ��
�� .sysoexecTEXT���     TEXT�� �j  � �� 0���������� 0 
openfolder 
OpenFolder�� ����� �  ���� 0 
folderpath 
FolderPath��  � ���� 0 
folderpath 
FolderPath�  8 :��
�� .sysoexecTEXT���     TEXT�� 
�%�%j � �� A���������� *0 getfolderfromdialog GetFolderFromDialog��  ��  �  � 	�� K��������������
�� 
prmp
�� 
dflc
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp
�� 
TEXT�� 
�� .sysostflalis    ��� null�� *����j �,�&� �,�&� �� Z���������� &0 getfilefromdialog GetFileFromDialog��  ��  �  � 	�� d��������������
�� 
prmp
�� 
dflc
�� afdrcusr
�� .earsffdralis        afdr
�� 
psxp
�� 
TEXT�� 
�� .sysostdfalis    ��� null�� *����j �,�&� �,�&� �� s���������� .0 getsubfoldersinfolder GetSubFoldersInFolder�� ����� �  ���� 0 
folderpath 
FolderPath��  � ���������� 0 
folderpath 
FolderPath�� 0 r  �� 0 	myfolders 	myFolders�� 0 f  �  � }���������� �
�� 
cfol
�� 
kocl
�� 
cobj
�� .corecnte****       ****
�� 
TEXT�� 4� 0�E�O*�/�-E�O �[��l kh ��%�&�%E�[OY��O�U� �� ����������� $0 getfilesinfolder GetFilesInFolder�� ��� �  �~�~ 0 	posixpath 	POSIXPath��  � �}�} 0 	posixpath 	POSIXPath�  � ��|
�| .sysoexecTEXT���     TEXT�� 
�%�%j � �{ ��z�y���x�{ 0 killfileonmac KillFileOnMac�z �w��w �  �v�v 0 filename FileName�y  � �u�u 0 filename FileName�  � ��t�s�r
�t 
psxp
�s 
strq
�r .sysoexecTEXT���     TEXT�x � ��,�,%j U� �q ��p�o���n�q "0 killfolderonmac KillFolderOnMac�p �m��m �  �l�l 0 
foldername 
FolderName�o  � �k�k 0 
foldername 
FolderName�  � ��j�i�h
�j 
psxp
�i 
strq
�h .sysoexecTEXT���     TEXT�n � ��,�,%j U� �g ��f�e���d�g 0 startrecord StartRecord�f �c��c �  �b�b 0 paramstring paramString�e  � �a�a 0 paramstring paramString�  ��`�_ ��^
�` .MVWRnwarnull��� ��� null
�_ 
docu
�^ .MVWRstarnull���     docu�d � *j O*��/j UOk� �] ��\�[���Z�] 0 
saverecord 
SaveRecord�\ �Y��Y �  �X�X 0 filename FileName�[  � �W�V�U�W 0 filename FileName�V 0 
exportpath 
exportPath�U 0 doc  � �T$�S�R��Q�P�O�N�M�L�K�J�I�H�G
�T 
psxf
�S 
docu
�R .MVWRstopnull���     docu�  
�Q 
pnam
�P 
cobj
�O 
kfil
�N 
file
�M 
expp�L 
�K .MVWRexponull���     docu
�J 
savo
�I savono  
�H .coreclosnull���     obj 
�G .aevtquitnull��� ��� null�Z G*�/E�O� 9*��/j O*�-�[�,\Z�@1�i/E�O��*�/��� O��a l O*j UOk� �F.�E�D���C�F 20 gethorizontalwindowsize GetHorizontalWindowSize�E  �D  � �B�A�@�B *0 visiblescreenbounds visibleScreenBounds�A 0 leftpos leftPos�@ 	0 width  � �?�>�=�<�;�:a
�? misccura�> 0 nsscreen NSScreen�= 0 
mainscreen 
mainScreen�< 0 visibleframe visibleFrame
�; 
cobj
�: 
TEXT�C -��,j+ j+ E�O��k/�k/E�O��l/�k/E�O��&�%�%�&� �9h�8�7���6�9 0 splitstring SplitString�8 �5��5 �  �4�3�4 0 thebigstring TheBigString�3  0 fieldseparator fieldSeparator�7  � �2�1�0�/�2 0 thebigstring TheBigString�1  0 fieldseparator fieldSeparator�0 0 oldtid oldTID�/ 0 theitems theItems� �.�-�,
�. 
ascr
�- 
txdl
�, 
citm�6  � *�,E�O�*�,FO��-E�O�*�,FUO�� �+��*�)���(�+ 0 activatetimer ActivateTimer�* �'��' �  �&�& 0 timerapp TimerApp�)  � �%�% 0 timerapp TimerApp� �$�#
�$ 
capp
�# .miscactvnull��� ��� null�( *�/ *j U� �"��!� ����" (0 getfromcitecreator GetFromCiteCreator�! ��� �  �� 0 paramstring paramString�   � �� 0 paramstring paramString� 
����������
� .miscactvnull��� ��� null
� .sysodelanull��� ��� nmbr
� 
faal
� eMdsKctl
� eMdsKopt
� .prcskprsnull���     ctxt� 2� #*j Okj O� ����lvl UOkj UO� *j U ascr  ��ޭ