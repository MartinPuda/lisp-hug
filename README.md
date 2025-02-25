# lisp-hug

My contributions to lisp-hug@lispworks.com (mailing list for LispWorks users).

## Available font weights

Prints all available font weights:
`(:BLACK :BOLD :DEMIBOLD :EXTRA-BOLD :EXTRA-LIGHT :LIGHT :MEDIUM :NORMAL)` 

```
(loop with weights = '()
      for font in (gp:find-matching-fonts (capi:create-dummy-graphics-port)
                                          (gp:make-font-description))
      do (pushnew (getf (gp:font-description-attributes (gp:font-description font)) :weight) weights)
      finally (print (sort weights #'string<)))
```
