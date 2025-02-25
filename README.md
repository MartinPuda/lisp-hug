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

## Excel Automation

Additional code for example file `(example-edit-file "com/automation/excel/pie-chart")`.

### Save chart
Add under `(excel-chart-wizard chart ...)`:
```
(multiple-value-bind (filename successp filter-name)
    (capi:prompt-for-file "Save Chart" :operation :save)
  (declare (ignore filter-name))
  (when successp
    (com:invoke-dispatch-method chart "Export" (string-append (namestring filename) ".png") "PNG")))
```

### Set color or image of given chart segment
Based on user's code in Visual Basic.
Note: Excel uses BBGGRR.
Note 2: [com:do-collection-items](https://www.lispworks.com/documentation/lw50/COM/html/com-116.htm) can be useful.
```
(defun method-chain (o &rest methods)
  (reduce (lambda (o method)
            (if (listp method)
                (apply #'com:invoke-dispatch-method o method)
              (com:invoke-dispatch-method o method)))
          methods
          :initial-value o))

(defun rgb-to-int (rgb)
  (+ (ash (third rgb) 16) 
     (ash (second rgb) 8) 
     (first rgb)))

(defun my-pix (chart)
  (let ((pts (method-chain chart "SeriesCollection" '("Item" 1) "Points")))
    (com:invoke-dispatch-put-property 
     (method-chain pts '("Item" 1) "Format" "Fill" "ForeColor")
     "RGB"
     (rgb-to-int '(255 255 0)))
    (method-chain pts '("Item" 2) "Format" "Fill" '("UserPicture" "path/to/file.png"))))
```
