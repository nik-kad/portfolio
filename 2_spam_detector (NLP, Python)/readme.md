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

Then I used another my class `Categorizator`, which helps to categorize text data by the similarity. The method `cat_sim` calculates similarity between the specified phrase and text_data and sorts them by descending the similarity.  
<p align="center">
<img src="./pictures/2024-08-08_21-13-08 (2).png" width=700>
</p>

This way almost all spam got located in the top of the table and it helped me to label this data easily and prepare the training dataset.  
To estimate this labeling I calculated some metrics:
| Metric | Value |
|:--|----:|
| Accuracy| 98.7% |
| F1 | 80.9% |

Quite a good metrics, but not perfect.  

Next I used the special method `get_train_data` of the class TextPreprocessing to prepare train data.  
For the model training I took pretrained Spacy model for Russian language and fine-tuned it using obtained train data.  
After this I tested the fine-tuned model both source and new data. As the new data I used a set of the generated texts which are very similar to the source texts, but the model had never seen them before.
|Type of Test Data| Metric | Value |
|:---|:--|----:|
|Source| Accuracy| 99.5% |
|Source| F1 | 93.8% |
|New | Accuracy |66.3%   |

The new test data consists of only spam-messages, so `f1`-metric didn't make sense and I calculated only `accuracy`.  
Unfortunately, the result wasn't good and this reported us about a low generalization level of the model in the task of spam detection.  

<p align="center">
Spam-detection result on the unknown data (before enriching)
<img src="./pictures/2024-08-09_14-45-01.png" width=700>
</p>  

Obviously, that the reason for this problem was connected with the small quantity of spam examples (only 3.8%)  

To solve this problem I used ChatGPT and YandexGPT to generate new spam-messages similar to the source spam-messages. It was a kind of text augmentation.  
For this purpose I wrote the special script to connect with the API of these LLM.
After enriching we obtained about 20% spam-examples.  
Retraining the model on the enriched dataset showed the significantly better result.  

<p align="center">
Spam-detection result on the unknown data (after enriching)
<img src="./pictures/2024-08-09_14-59-39.png" width=700>
</p>  

| Metric | Value |
|:--|----:|
| Accuracy| 99.4% |

The metric is good and the generalization of this model is significantly better than the generalization of the previous one.  

Spam filter for this chat is ready.

See the full [notebook](spam_detection3p.ipynb).



