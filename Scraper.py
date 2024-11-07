from components import *

result = pdf_extractor(count=1000, newExtraction=False)
if type(result) == list:
    result.insert(5, "1511.9030")
    print(result)
    run(result)
else:
    print(result)