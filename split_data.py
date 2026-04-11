 i m p o r t   o s 
 i m p o r t   p a n d a s   a s   p d 
 
 d a t a _ d i r   =   ' d a t a ' 
 o u t _ d i r   =   ' d a t a 2 0 0 ' 
 c h u n k _ s i z e   =   2 0 0 
 
 t o t a l _ f i l e s _ p r o c e s s e d   =   0 
 t o t a l _ c h u n k s _ c r e a t e d   =   0 
 
 f o r   r o o t ,   d i r s ,   f i l e s   i n   o s . w a l k ( d a t a _ d i r ) : 
         f o r   f   i n   f i l e s : 
                 i f   f . e n d s w i t h ( ' . c s v ' ) : 
                         f i l e p a t h   =   o s . p a t h . j o i n ( r o o t ,   f ) 
                         p r i n t ( f " P r o c e s s i n g   { f i l e p a t h } . . . " ) 
                         
                         t r y : 
                                 d f   =   p d . r e a d _ c s v ( f i l e p a t h ,   h e a d e r = N o n e ) 
                         e x c e p t   E x c e p t i o n   a s   e : 
                                 p r i n t ( f " E r r o r   r e a d i n g   { f i l e p a t h } :   { e } " ) 
                                 c o n t i n u e 
                                 
                         n u m _ c h u n k s   =   l e n ( d f )   / /   c h u n k _ s i z e 
                         i f   n u m _ c h u n k s   = =   0 : 
                                 p r i n t ( f "     - >   S k i p p i n g ,   n o t   e n o u g h   d a t a   ( { l e n ( d f ) }   r o w s ) . " ) 
                                 c o n t i n u e 
                                 
                         b a s e _ n a m e   =   o s . p a t h . s p l i t e x t ( f ) [ 0 ] 
                         
                         #   C r e a t e   c o r r e s p o n d i n g   s u b d i r e c t o r i e s   i n   o u t _ d i r 
                         r e l _ p a t h   =   o s . p a t h . r e l p a t h ( r o o t ,   d a t a _ d i r ) 
                         t a r g e t _ r o o t   =   o s . p a t h . j o i n ( o u t _ d i r ,   r e l _ p a t h ) 
                         o s . m a k e d i r s ( t a r g e t _ r o o t ,   e x i s t _ o k = T r u e ) 
                         
                         f o r   i   i n   r a n g e ( n u m _ c h u n k s ) : 
                                 c h u n k   =   d f . i l o c [ i * c h u n k _ s i z e   :   ( i + 1 ) * c h u n k _ s i z e ] 
                                 o u t _ n a m e   =   f " { b a s e _ n a m e } _ p a r t { i + 1 } . c s v " 
                                 o u t _ p a t h   =   o s . p a t h . j o i n ( t a r g e t _ r o o t ,   o u t _ n a m e ) 
                                 c h u n k . t o _ c s v ( o u t _ p a t h ,   h e a d e r = F a l s e ,   i n d e x = F a l s e ) 
                                 t o t a l _ c h u n k s _ c r e a t e d   + =   1 
                                 
                         p r i n t ( f "     - >   C r e a t e d   { n u m _ c h u n k s }   c h u n k s   i n   { t a r g e t _ r o o t }   ( r e m a i n i n g   { l e n ( d f )   %   c h u n k _ s i z e }   r o w s   d i s c a r d e d ) . " ) 
                         t o t a l _ f i l e s _ p r o c e s s e d   + =   1 
 
 p r i n t ( f " \ n D o n e !   P r o c e s s e d   { t o t a l _ f i l e s _ p r o c e s s e d }   o r i g i n a l   f i l e s ,   c r e a t e d   { t o t a l _ c h u n k s _ c r e a t e d }   c h u n k e d   f i l e s . " ) 
 