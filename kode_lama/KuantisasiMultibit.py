 # ! / u s r / b i n / e n v   p y t h o n 3 
 " " " 
 K u a n t i s a s i   M u l t i b i t   ( F i x e d   /   N o n - A d a p t i v e ,   G r a y - C o d e d ) 
 = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 -   M e m b a c a   d u a   f i l e   C S V   ( A l i c e   &   B o b ) 
 -   M e n g k u a n t i s a s i   d a t a   d e n g a n   j u m l a h   b i t   t e t a p   ( u n i f o r m   q u a n t i z a t i o n ) 
 -   M e n g h a s i l k a n   b i t s t r e a m   G r a y   c o d e   u n t u k   t i a p   d a t a   p o i n t 
 -   M e n g h i t u n g   e n t r o p i ,   K G R ,   K D R 
 -   M e n y i m p a n   h a s i l   k e   C S V   d a n   E x c e l 
 " " " 
 
 i m p o r t   o s 
 i m p o r t   m a t h 
 i m p o r t   t i m e 
 i m p o r t   c s v 
 i m p o r t   x l w t 
 i m p o r t   p a n d a s   a s   p d 
 i m p o r t   n u m p y   a s   n p 
 f r o m   t y p i n g   i m p o r t   L i s t 
 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 #   U t i l i t y   f u n c t i o n s 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 d e f   g r a y _ c o d e ( n :   i n t )   - >   L i s t [ s t r ] : 
         " " " G e n e r a t e   n - b i t   G r a y   c o d e   s e q u e n c e . " " " 
         i f   n   < =   0 : 
                 r e t u r n   [ ' 0 ' ] 
         c o d e s   =   [ ' 0 ' ,   ' 1 ' ] 
         f o r   b i t s   i n   r a n g e ( 2 ,   n + 1 ) : 
                 m i r r o r   =   c o d e s [ : : - 1 ] 
                 c o d e s   =   [ ' 0 '   +   c   f o r   c   i n   c o d e s ]   +   [ ' 1 '   +   c   f o r   c   i n   m i r r o r ] 
         r e t u r n   c o d e s 
 
 d e f   c a l c u l a t e _ e n t r o p y ( b i t s t r e a m :   s t r )   - >   f l o a t : 
         " " " S h a n n o n   e n t r o p y   o f   b i t s t r e a m . " " " 
         i f   n o t   b i t s t r e a m : 
                 r e t u r n   0 . 0 
         p 0   =   b i t s t r e a m . c o u n t ( ' 0 ' )   /   l e n ( b i t s t r e a m ) 
         p 1   =   b i t s t r e a m . c o u n t ( ' 1 ' )   /   l e n ( b i t s t r e a m ) 
         e n t   =   0 . 0 
         i f   p 0   >   0 : 
                 e n t   - =   p 0   *   m a t h . l o g 2 ( p 0 ) 
         i f   p 1   >   0 : 
                 e n t   - =   p 1   *   m a t h . l o g 2 ( p 1 ) 
         r e t u r n   e n t 
 
 d e f   c a l c u l a t e _ k d r ( a :   s t r ,   b :   s t r )   - >   f l o a t : 
         " " " K e y   D i s a g r e e m e n t   R a t e   ( % ) " " " 
         i f   n o t   a   o r   n o t   b : 
                 r e t u r n   0 . 0 
         n   =   m i n ( l e n ( a ) ,   l e n ( b ) ) 
         i f   n   = =   0 : 
                 r e t u r n   0 . 0 
         d i f f   =   s u m ( 1   f o r   i   i n   r a n g e ( n )   i f   a [ i ]   ! =   b [ i ] ) 
         r e t u r n   ( d i f f   /   n )   *   1 0 0 . 0 
 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 #   K u a n t i s a s i   M u l t i b i t   ( F i x e d   /   N o n - A d a p t i v e ) 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 d e f   m u l t i b i t _ q u a n t i z a t i o n ( 
         d a t a :   n p . n d a r r a y , 
         n u m _ b i t s :   i n t   =   3 
 ) : 
         " " " 
         M e l a k u k a n   k u a n t i s a s i   m u l t i b i t   t e t a p   ( u n i f o r m )   d a n   m e n g h a s i l k a n   b i t s t r e a m   G r a y   c o d e . 
         " " " 
         t 0   =   t i m e . p e r f _ c o u n t e r ( ) 
         d a t a   =   n p . a s a r r a y ( d a t a ) . a s t y p e ( f l o a t ) 
         n   =   l e n ( d a t a ) 
         i f   n   = =   0 : 
                 r e t u r n   " " ,   0 ,   0 . 0 ,   0 . 0 ,   0 . 0 
 
         n u m _ b i t s   =   i n t ( m a x ( 1 ,   m i n ( 8 ,   n u m _ b i t s ) ) ) 
         l e v e l s   =   2   * *   n u m _ b i t s 
 
         d a t a _ m i n   =   n p . m i n ( d a t a ) 
         d a t a _ m a x   =   n p . m a x ( d a t a ) 
         d a t a _ r a n g e   =   d a t a _ m a x   -   d a t a _ m i n 
 
         i f   d a t a _ r a n g e   = =   0 : 
                 #   S e m u a   n i l a i   s a m a   →   m a p   k e   s a t u   l e v e l   ( 0 ) 
                 i n d i c e s   =   n p . z e r o s ( n ,   d t y p e = i n t ) 
         e l s e : 
                 #   U n i f o r m   q u a n t i z a t i o n 
                 s t e p   =   d a t a _ r a n g e   /   l e v e l s 
                 i n d i c e s   =   n p . f l o o r ( ( d a t a   -   d a t a _ m i n )   /   s t e p ) . a s t y p e ( i n t ) 
                 i n d i c e s [ i n d i c e s   > =   l e v e l s ]   =   l e v e l s   -   1 
 
         g r a y _ m a p   =   g r a y _ c o d e ( n u m _ b i t s ) 
         b i t _ l i s t   =   [ g r a y _ m a p [ i ]   f o r   i   i n   i n d i c e s ] 
         b i t s t r e a m   =   " " . j o i n ( b i t _ l i s t ) 
 
         t o t a l _ b i t s   =   l e n ( b i t s t r e a m ) 
         e n t r o p y   =   c a l c u l a t e _ e n t r o p y ( b i t s t r e a m ) 
 
         t 1   =   t i m e . p e r f _ c o u n t e r ( ) 
         e l a p s e d   =   t 1   -   t 0 
         i f   e l a p s e d   = =   0 : 
                 e l a p s e d   =   1 e - 9 
         k g r   =   t o t a l _ b i t s   /   e l a p s e d     #   b i t / s 
 
         r e t u r n   b i t s t r e a m ,   t o t a l _ b i t s ,   e n t r o p y ,   k g r ,   e l a p s e d 
 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 #   B e n c h m a r k 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 d e f   b e n c h m a r k _ k u a n t i s a s i ( s e r i e s :   p d . S e r i e s ,   n u m _ b i t s :   i n t   =   3 ,   r u n s :   i n t   =   1 0 ) : 
         d a t a   =   s e r i e s . d r o p n a ( ) . v a l u e s . a s t y p e ( f l o a t ) 
         b i t s t r e a m s ,   t i m e s ,   e n t r o p i e s ,   k g r s ,   l e n g t h s   =   [ ] ,   [ ] ,   [ ] ,   [ ] ,   [ ] 
 
         f o r   i   i n   r a n g e ( r u n s ) : 
                 b i t s t r e a m ,   t o t a l _ b i t s ,   e n t r o p y ,   k g r ,   t   =   m u l t i b i t _ q u a n t i z a t i o n ( d a t a ,   n u m _ b i t s = n u m _ b i t s ) 
                 p r i n t ( f " [ R u n   { i + 1 } ]   b i t s = { t o t a l _ b i t s } ,   e n t r o p y = { e n t r o p y : . 4 f } ,   t i m e = { t : . 6 f } s ,   K G R = { k g r : . 2 f } " ) 
                 b i t s t r e a m s . a p p e n d ( b i t s t r e a m ) 
                 t i m e s . a p p e n d ( t ) 
                 e n t r o p i e s . a p p e n d ( e n t r o p y ) 
                 k g r s . a p p e n d ( k g r ) 
                 l e n g t h s . a p p e n d ( t o t a l _ b i t s ) 
 
         r e t u r n   b i t s t r e a m s ,   t i m e s ,   e n t r o p i e s ,   k g r s ,   l e n g t h s 
 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 #   O u t p u t   u t i l i t i e s 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 d e f   s a v e _ b i t s t r e a m _ t o _ c s v ( b i t s t r e a m :   s t r ,   p a t h :   s t r ) : 
         o s . m a k e d i r s ( o s . p a t h . d i r n a m e ( p a t h ) ,   e x i s t _ o k = T r u e ) 
         w i t h   o p e n ( p a t h ,   " w " ,   n e w l i n e = ' ' )   a s   f : 
                 w r i t e r   =   c s v . w r i t e r ( f ) 
                 w r i t e r . w r i t e r o w ( [ " b i t s t r e a m " ] ) 
                 w r i t e r . w r i t e r o w ( [ b i t s t r e a m ] ) 
         p r i n t ( f " [ O K ]   B i t s t r e a m   d i s i m p a n   k e :   { p a t h } " ) 
 
 d e f   s a v e _ k d r _ k g r _ e x c e l ( 
         b i t s _ a ,   b i t s _ b , 
         k g r s _ a ,   k g r s _ b , 
         t i m e s _ a ,   t i m e s _ b , 
         o u t p a t h = " O u t p u t / P 2 P / h a s i l k u a n t i s a s i / a n a l i s i s _ r i n g k a s _ e v e . x l s " 
 ) : 
         o s . m a k e d i r s ( o s . p a t h . d i r n a m e ( o u t p a t h ) ,   e x i s t _ o k = T r u e ) 
         b o o k   =   x l w t . W o r k b o o k ( ) 
         s h e e t   =   b o o k . a d d _ s h e e t ( " K D R - K G R " ) 
 
         h e a d e r s   =   [ " P e r c o b a a n " ,   " K D R   ( % ) " ,   " K G R   A l i c e " ,   " T i m e   A l i c e   ( s ) " ,   " K G R   B o b " ,   " T i m e   B o b   ( s ) " ] 
         f o r   j ,   h   i n   e n u m e r a t e ( h e a d e r s ) : 
                 s h e e t . w r i t e ( 0 ,   j ,   h ) 
 
         r u n s   =   m i n ( l e n ( b i t s _ a ) ,   l e n ( b i t s _ b ) ) 
         f o r   i   i n   r a n g e ( r u n s ) : 
                 k d r   =   c a l c u l a t e _ k d r ( b i t s _ a [ i ] ,   b i t s _ b [ i ] ) 
                 s h e e t . w r i t e ( i + 1 ,   0 ,   f " R u n   { i + 1 } " ) 
                 s h e e t . w r i t e ( i + 1 ,   1 ,   r o u n d ( k d r ,   3 ) ) 
                 s h e e t . w r i t e ( i + 1 ,   2 ,   r o u n d ( k g r s _ a [ i ] ,   2 ) ) 
                 s h e e t . w r i t e ( i + 1 ,   3 ,   r o u n d ( t i m e s _ a [ i ] ,   6 ) ) 
                 s h e e t . w r i t e ( i + 1 ,   4 ,   r o u n d ( k g r s _ b [ i ] ,   2 ) ) 
                 s h e e t . w r i t e ( i + 1 ,   5 ,   r o u n d ( t i m e s _ b [ i ] ,   6 ) ) 
 
         b o o k . s a v e ( o u t p a t h ) 
         p r i n t ( f " 📊   F i l e   a n a l i s i s   d i s i m p a n   k e :   { o u t p a t h } " ) 
 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 #   M a i n   C L I 
 #   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
 
 d e f   r e a d _ f i r s t _ c o l u m n _ c s v ( p a t h :   s t r )   - >   p d . S e r i e s : 
         " " " A m b i l   k o l o m   p e r t a m a   d a r i   C S V   ( h e a d e r   o p s i o n a l ) . " " " 
         t r y : 
                 d f   =   p d . r e a d _ c s v ( p a t h ,   h e a d e r = 0 ) 
                 c o l   =   d f . c o l u m n s [ 0 ] 
                 s e r i e s   =   p d . t o _ n u m e r i c ( d f [ c o l ] ,   e r r o r s = ' c o e r c e ' ) . d r o p n a ( ) 
         e x c e p t   E x c e p t i o n : 
                 d f   =   p d . r e a d _ c s v ( p a t h ,   h e a d e r = N o n e ) 
                 s e r i e s   =   p d . t o _ n u m e r i c ( d f . i l o c [ : , 0 ] ,   e r r o r s = ' c o e r c e ' ) . d r o p n a ( ) 
         r e t u r n   s e r i e s 
 
 d e f   m a i n ( ) : 
         p r i n t ( " = = =   K U A N T I S A S I   M U L T I B I T   ( N O N - A D A P T I V E )   = = = \ n " ) 
 
         a l i c e _ p a t h   =   i n p u t ( " M a s u k k a n   p a t h   f i l e   C S V   u n t u k   A l i c e :   " ) . s t r i p ( ) 
         b o b _ p a t h       =   i n p u t ( " M a s u k k a n   p a t h   f i l e   C S V   u n t u k   B o b     :   " ) . s t r i p ( ) 
 
         i f   n o t   o s . p a t h . e x i s t s ( a l i c e _ p a t h )   o r   n o t   o s . p a t h . e x i s t s ( b o b _ p a t h ) : 
                 p r i n t ( " [ E R R O R ]   F i l e   t i d a k   d i t e m u k a n . " ) 
                 r e t u r n 
 
         d f _ a   =   r e a d _ f i r s t _ c o l u m n _ c s v ( a l i c e _ p a t h ) 
         d f _ b   =   r e a d _ f i r s t _ c o l u m n _ c s v ( b o b _ p a t h ) 
 
         n u m _ b i t s   =   i n p u t ( " M a s u k k a n   j u m l a h   b i t   p e r   s a m p l e   ( 1 - 8 ) :   " ) . s t r i p ( ) 
         i f   n o t   n u m _ b i t s . i s d i g i t ( ) : 
                 n u m _ b i t s   =   3 
         e l s e : 
                 n u m _ b i t s   =   i n t ( n u m _ b i t s ) 
         r u n s   =   i n p u t ( " M a s u k k a n   j u m l a h   p e r c o b a a n   ( d e f a u l t   1 0 ) :   " ) . s t r i p ( ) 
         r u n s   =   i n t ( r u n s )   i f   r u n s . i s d i g i t ( )   a n d   i n t ( r u n s )   >   0   e l s e   1 0 
 
         o u t _ d i r   =   " O u t p u t / P 2 P / h a s i l k u a n t i s a s i _ e v e " 
         o s . m a k e d i r s ( o u t _ d i r ,   e x i s t _ o k = T r u e ) 
 
         p r i n t ( f " \ n - - -   A l i c e   ( n u m _ b i t s = { n u m _ b i t s } )   - - - " ) 
         b i t s _ a ,   t i m e s _ a ,   e n t _ a ,   k g r _ a ,   l e n s _ a   =   b e n c h m a r k _ k u a n t i s a s i ( d f _ a ,   n u m _ b i t s = n u m _ b i t s ,   r u n s = r u n s ) 
         s a v e _ b i t s t r e a m _ t o _ c s v ( b i t s _ a [ - 1 ] ,   o s . p a t h . j o i n ( o u t _ d i r ,   " a l i c e _ b i t s t r e a m . c s v " ) ) 
 
         p r i n t ( f " \ n - - -   B o b   ( n u m _ b i t s = { n u m _ b i t s } )   - - - " ) 
         b i t s _ b ,   t i m e s _ b ,   e n t _ b ,   k g r _ b ,   l e n s _ b   =   b e n c h m a r k _ k u a n t i s a s i ( d f _ b ,   n u m _ b i t s = n u m _ b i t s ,   r u n s = r u n s ) 
         s a v e _ b i t s t r e a m _ t o _ c s v ( b i t s _ b [ - 1 ] ,   o s . p a t h . j o i n ( o u t _ d i r ,   " b o b _ b i t s t r e a m . c s v " ) ) 
 
         s a v e _ k d r _ k g r _ e x c e l ( b i t s _ a ,   b i t s _ b ,   k g r _ a ,   k g r _ b ,   t i m e s _ a ,   t i m e s _ b , 
                                               o s . p a t h . j o i n ( o u t _ d i r ,   " a n a l i s i s _ r i n g k a s _ e v e . x l s " ) ) 
 
         p r i n t ( " \ n [ O K ]   P r o s e s   s e l e s a i !   S e m u a   h a s i l   a d a   d i   f o l d e r : " ) 
         p r i n t ( o u t _ d i r ) 
 
 i f   _ _ n a m e _ _   = =   " _ _ m a i n _ _ " : 
         m a i n ( ) 
 