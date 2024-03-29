<p>
  <img width=200" align='left' src="https://github.com/nik-kad/portfolio/blob/main/pictures/text_classification.jpg">
</p>

### NLP Text Classification and Word Classification


The main goal of this project was an information extraction from raw unstructured text.

I got the dataset with the columns which contained an information about clients' career and interests as unstructured text with a lot of unnecessary information.  
Regarding the career I needed to remain only words meant name of professions.

<p>
  <img align='center' src="https://github.com/nik-kad/portfolio/blob/main/pictures/career_extr_dark.png">
</p>

In the case of interests I had to categorize the texts in the 50 specified categories, for example: 'Astronomy and space', 'Sports and fitness', 'Stocks, investment opportunities, investing money', etc.

<p>
  <img align='center' src="https://github.com/nik-kad/portfolio/blob/main/pictures/textcat.png">
</p>

To reach these goals I used Spacy and wrote several useful functions in Python.

[see the code](https://github.com/nik-kad/portfolio/blob/main/WORK%20PROJECTS/career%26interests_extraction/nk_nlp1_3.py)

I prepared learning data and fine-tuned pretrained models for English and Russian to recognize the words related to professions.
I also added the layer of text categorization into the models and trained it.
The trained models show a good ability to generalize.

[Learn more...](https://github.com/nik-kad/portfolio/blob/main/pictures/text_classification.jpg)

---

