 i m p o r t   n u m p y   a s   n p 
 i m p o r t   t i m e 
 i m p o r t   c s v 
 i m p o r t   o s 
 f r o m   o p e n p y x l   i m p o r t   W o r k b o o k 
 f r o m   s c i p y . s t a t s   i m p o r t   p e a r s o n r 
 
 #   = = =   P a r a m e t e r   K a l m a n   = = = 
 a   =   1 
 h   =   1 
 R   =   1 
 Q   =   0 . 0 1 
 x a p o s t e r i o r i _ 0   =   - 5 
 p a p o s t e r i o r i _ 0   =   1 
 b b   =   6     #   j u m l a h   m e a s u r e m e n t   p e r   s a m p l e 
 i n t e r v a l   =   0 . 1 1 
 
 #   = = =   B a c a   d a t a   d a r i   f i l e   C S V   t e r p i s a h   = = = 
 d e f   r e a d _ r s s i _ c s v ( p a t h ) : 
         d a t a   =   [ ] 
         w i t h   o p e n ( p a t h ,   ' r ' )   a s   f : 
                 r e a d e r   =   c s v . r e a d e r ( f ) 
                 f o r   r o w   i n   r e a d e r : 
                         t r y : 
                                 d a t a . a p p e n d ( i n t ( r o w [ 0 ] ) )     #   b a c a   k o l o m   p e r t a m a 
                         e x c e p t : 
                                 c o n t i n u e     #   l e w a t i   b a r i s   k o s o n g   a t a u   h e a d e r 
         r e t u r n   d a t a 
 
 r s s _ a l i c e   =   r e a d _ r s s i _ c s v ( ' s k e n a r i o 1 _ m i t a _ e v e a l i c e . c s v ' ) 
 r s s _ b o b       =   r e a d _ r s s i _ c s v ( ' s k e n a r i o 1 _ m i t a _ e v e b o b . c s v ' ) 
 
 #   = = =   P e r s i a p a n   r e s h a p i n g   = = = 
 t o t a l _ d a t a   =   m i n ( l e n ( r s s _ a l i c e ) ,   l e n ( r s s _ b o b ) )     #   p a s t i k a n   p a n j a n g   s a m a 
 a a   =   t o t a l _ d a t a   / /   b b 
 r s s _ a l i c e   =   r s s _ a l i c e [ : a a   *   b b ] 
 r s s _ b o b   =   r s s _ b o b [ : a a   *   b b ] 
 a l i c e   =   n p . a r r a y ( r s s _ a l i c e ) . r e s h a p e ( a a ,   b b ) . T 
 b o b   =   n p . a r r a y ( r s s _ b o b ) . r e s h a p e ( a a ,   b b ) . T 
 
 #   = = =   F u n g s i   K a l m a n   F i l t e r   = = = 
 d e f   k a l m a n _ f i l t e r ( s i g n a l ) : 
         x a p o s t e r i o r i   =   [ ] 
         p a p o s t e r i o r i   =   [ ] 
         r o w 1 = [ ] ;   r o w 2 = [ ] ;   r o w 3 = [ ] ;   r o w 4 = [ ] ;   r o w 5 = [ ] ;   r o w 6 = [ ] 
 
         f o r   m   i n   r a n g e ( a a ) : 
                 r o w 1 . a p p e n d ( a   *   x a p o s t e r i o r i _ 0 ) 
                 r o w 2 . a p p e n d ( s i g n a l [ 0 ] [ m ]   -   h   *   r o w 1 [ m ] ) 
                 r o w 3 . a p p e n d ( a * a   *   p a p o s t e r i o r i _ 0   +   Q ) 
                 g a i n   =   r o w 3 [ m ]   /   ( r o w 3 [ m ]   +   R ) 
                 r o w 4 . a p p e n d ( g a i n ) 
                 r o w 5 . a p p e n d ( r o w 3 [ m ]   *   ( 1   -   g a i n ) ) 
                 r o w 6 . a p p e n d ( r o w 1 [ m ]   +   g a i n   *   r o w 2 [ m ] ) 
         x a p o s t e r i o r i . a p p e n d ( r o w 6 ) 
         p a p o s t e r i o r i . a p p e n d ( r o w 5 ) 
 
         f o r   j   i n   r a n g e ( 1 ,   b b ) : 
                 r 1 = [ ] ;   r 2 = [ ] ;   r 3 = [ ] ;   r 4 = [ ] ;   r 5 = [ ] ;   r 6 = [ ] 
                 f o r   n   i n   r a n g e ( a a ) : 
                         r 1 . a p p e n d ( x a p o s t e r i o r i [ j - 1 ] [ n ] ) 
                         r 2 . a p p e n d ( s i g n a l [ j ] [ n ]   -   h   *   r 1 [ n ] ) 
                         r 3 . a p p e n d ( a * a   *   p a p o s t e r i o r i [ j - 1 ] [ n ]   +   Q ) 
                         g a i n   =   r 3 [ n ]   /   ( r 3 [ n ]   +   R ) 
                         r 4 . a p p e n d ( g a i n ) 
                         r 5 . a p p e n d ( r 3 [ n ]   *   ( 1   -   g a i n ) ) 
                         r 6 . a p p e n d ( r 1 [ n ]   +   g a i n   *   r 2 [ n ] ) 
                 x a p o s t e r i o r i . a p p e n d ( r 6 ) 
                 p a p o s t e r i o r i . a p p e n d ( r 5 ) 
         r e t u r n   x a p o s t e r i o r i 
 
 #   = = =   B u a t   d i r e k t o r i   o u t p u t   = = = 
 o s . m a k e d i r s ( ' O u t p u t / P 2 P / h a s i l p r a p r o s e s ' ,   e x i s t _ o k = T r u e ) 
 
 #   = = =   S i m p a n   h a s i l   K a l m a n   A l i c e   = = = 
 h a s i l _ a l i c e   =   n p . a r r a y ( k a l m a n _ f i l t e r ( a l i c e ) ) . T . r e s h a p e ( - 1 ,   1 ) 
 w i t h   o p e n ( ' O u t p u t / P 2 P / h a s i l p r a p r o s e s / e v e _ a l i c e _ s k e n a r i o 1 . c s v ' ,   ' w ' ,   n e w l i n e = ' ' )   a s   f : 
         w r i t e r   =   c s v . w r i t e r ( f ) 
         w r i t e r . w r i t e r o w ( [ ' A l i c e _ p r a p r o s e s ' ] ) 
         f o r   v a l   i n   h a s i l _ a l i c e : 
                 w r i t e r . w r i t e r o w ( [ i n t ( v a l . i t e m ( ) ) ] ) 
 
 #   = = =   S i m p a n   h a s i l   K a l m a n   B o b   = = = 
 h a s i l _ b o b   =   n p . a r r a y ( k a l m a n _ f i l t e r ( b o b ) ) . T . r e s h a p e ( - 1 ,   1 ) 
 w i t h   o p e n ( ' O u t p u t / P 2 P / h a s i l p r a p r o s e s / e v e _ b o b _ s k e n a r i o 1 . c s v ' ,   ' w ' ,   n e w l i n e = ' ' )   a s   f : 
         w r i t e r   =   c s v . w r i t e r ( f ) 
         w r i t e r . w r i t e r o w ( [ ' B o b _ p r a p r o s e s ' ] ) 
         f o r   v a l   i n   h a s i l _ b o b : 
                 w r i t e r . w r i t e r o w ( [ i n t ( v a l . i t e m ( ) ) ] ) 
 
 #   = = =   K o r e l a s i   P e a r s o n   = = = 
 k o r e l a s i ,   _   =   p e a r s o n r ( h a s i l _ a l i c e . f l a t t e n ( ) ,   h a s i l _ b o b . f l a t t e n ( ) ) 
 
 #   = = =   B e n c h m a r k   &   K G R   = = = 
 d e f   b e n c h m a r k ( s i g n a l ,   l a b e l ) : 
         t i m e s   =   [ ] 
         k g r s   =   [ ] 
 
         f o r   i   i n   r a n g e ( 1 0 ) : 
                 s t a r t   =   t i m e . p e r f _ c o u n t e r ( ) 
                 r e s u l t   =   k a l m a n _ f i l t e r ( s i g n a l ) 
                 e n d   =   t i m e . p e r f _ c o u n t e r ( ) 
                 e l a p s e d   =   e n d   -   s t a r t 
                 g a i n _ c o u n t   =   a a   *   b b 
                 k g r   =   g a i n _ c o u n t   *   3 2   /   e l a p s e d 
                 p r i n t ( f " { l a b e l }   p e r c o b a a n   k e - { i + 1 } :   { e l a p s e d : . 6 f }   d e t i k   |   K G R :   { k g r : . 2 f }   b i t / s " ) 
                 t i m e s . a p p e n d ( e l a p s e d ) 
                 k g r s . a p p e n d ( k g r ) 
 
         a v g _ t i m e   =   s u m ( t i m e s )   /   l e n ( t i m e s ) 
         a v g _ k g r   =   s u m ( k g r s )   /   l e n ( k g r s ) 
         p r i n t ( f " R a t a - r a t a   w a k t u   K a l m a n   { l a b e l } :   { a v g _ t i m e : . 6 f }   d e t i k " ) 
         p r i n t ( f " R a t a - r a t a   K G R   K a l m a n   { l a b e l } :   { a v g _ k g r : . 2 f }   b i t / s \ n " ) 
         r e t u r n   t i m e s ,   k g r s ,   a v g _ t i m e ,   a v g _ k g r 
 
 #   = = =   J a l a n k a n   b e n c h m a r k   = = = 
 p r i n t ( " = = = = =   B E N C H M A R K   A L I C E   = = = = = " ) 
 t i m e s _ a l i c e ,   k g r _ a l i c e ,   a v g _ t i m e _ a l i c e ,   a v g _ k g r _ a l i c e   =   b e n c h m a r k ( a l i c e ,   ' A l i c e ' ) 
 
 p r i n t ( " = = = = =   B E N C H M A R K   B O B   = = = = = " ) 
 t i m e s _ b o b ,   k g r _ b o b ,   a v g _ t i m e _ b o b ,   a v g _ k g r _ b o b   =   b e n c h m a r k ( b o b ,   ' B o b ' ) 
 
 #   = = =   S i m p a n   h a s i l   a n a l i s i s   k e   E x c e l   = = = 
 w b   =   W o r k b o o k ( ) 
 w s   =   w b . a c t i v e 
 w s . t i t l e   =   " A n a l i s i s   K a l m a n   e v e " 
 
 w s . a p p e n d ( [ " I t e r a s i " ,   " W a k t u   A l i c e   ( s ) " ,   " K G R   A l i c e   ( b i t / s ) " ,   " W a k t u   B o b   ( s ) " ,   " K G R   B o b   ( b i t / s ) " ] ) 
 
 f o r   i   i n   r a n g e ( 1 0 ) : 
         w s . a p p e n d ( [ i   +   1 ,   t i m e s _ a l i c e [ i ] ,   k g r _ a l i c e [ i ] ,   t i m e s _ b o b [ i ] ,   k g r _ b o b [ i ] ] ) 
 
 f o r   r o w   i n   w s . i t e r _ r o w s ( m i n _ r o w = 2 ,   m a x _ r o w = 1 1 ,   m i n _ c o l = 2 ,   m a x _ c o l = 5 ) : 
         f o r   c e l l   i n   r o w : 
                 c e l l . n u m b e r _ f o r m a t   =   ' 0 . 0 0 0 0 0 0 ' 
 
 w s . a p p e n d ( [ " R A T A - R A T A " ,   a v g _ t i m e _ a l i c e ,   a v g _ k g r _ a l i c e ,   a v g _ t i m e _ b o b ,   a v g _ k g r _ b o b ] ) 
 w s . a p p e n d ( [ ] ) 
 w s . a p p e n d ( [ " K o r e l a s i   P e a r s o n " ,   k o r e l a s i ] ) 
 
 e x c e l _ p a t h   =   ' O u t p u t / P 2 P / h a s i l p r a p r o s e s / a n a l i s i s _ k a l m a n _ e v e . x l s x ' 
 w b . s a v e ( e x c e l _ p a t h ) 
 p r i n t ( f " H a s i l   a n a l i s i s   d i s i m p a n   k e   { e x c e l _ p a t h } " ) 
 