# odfdo-contrib

A few scripts using odfdo ODF library.

## odsgenerator.py

* generate .ods file from json description,
* description can be minimalist: a list of lists of lists,
* description can be complex, allowing styles at row or cell level,
* example:


        ./odsgenerator.py test/test_data_1.json test/report.ods
