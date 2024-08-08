## Spam detection using NLP

<img src="./pictures/photo_2024-08-08_15-40-28.jpg" width=300 align="left">This project consisted of two tasks:  

1) to find spam-messages in the set of texts I received from the client and remove them  

2) to build the model which will be able to detect spam-messages

To solve these tasks I used pretrained [Spacy](https://spacy.io/) model for Russian language and [my own module](./nk_nlp1_5.py) in Python `nk_nlp1_5`  

My module contains two classes:  

`TextPreprocessing` - this class has several methods which helps to process texts using such operations like: regular expressions, deduplication, mapping, quoting and NLP-methods based on semantic similarity, finding part-of-speech and sentence dependences, named entity recognition and allows applying these methods to the collection of texts directly  

`Categorizator` - the class whose methods can identify dependencies between a sets of texts. It supports several approaches to the similarity calculation which can be specified by the special parameters. In addition to the similarity the methods allows you to calculate quoting - how often this word or expression occurs in the other set of texts. This can help you to concentrate your attention on the most significant objects.

See help(classname) for details.

The pipeline of all the process looks so:
<img src="./pictures/pipeline.png" width=1000 align="left">  


After loading and preprocessing the source data, I used my class `TextPreprocessing` for data processing.  
Method `.word_extractor` is used to remain the specified number of the words, which are closest to the set of words or phrases given as argument `pattern`.  
We can also put some additional parameters into this method to filter words more finely. In this case I used filtering by *parts-of-speech*.  


