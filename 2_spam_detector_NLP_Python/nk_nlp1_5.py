### version 1.5

import pandas as pd
import numpy as np
import spacy
from spacy.matcher import Matcher
from spacy.tokens import Span, Doc, DocBin

import re
import os
import sys
from IPython.display import clear_output

from sklearn.model_selection import train_test_split
from sklearn.metrics.pairwise import cosine_similarity, euclidean_distances

import time

import ipywidgets as widgets
from ipywidgets import IntProgress, Label
from IPython.display import display

# defining a class for nlp-textprocessing

class TextPreprocessing:
    """ Process text data in the given column of dataset or series or list(tuple) using re, pandas and spacy libs.
    Using the specified methods you can clean text from unused data, extract required data, filter,
    get unique words and obtain additional statistic information about given text.
                    
    Parameters
    ----------
    text_col : list, tuple, pd.Series
        A list of text data used as an object for the processing.
    nlp : spacy model class
        Model for the specified language used for process.

    Attributes
    ----------
    nlp : spacy model class
        Model for the specified language using for process.
    text_col : pd.Series
        A list of text data used as an object for the processing.
    textcol_mod : pd.Series
        A copy of text data used for saving intermediate results.
    unique_tokens : pd.Series
        A list of unique tokens extracted from source texts
    LABELS_LIST_RU : list
        The list of named entities for Russian language
    LABELS_LIST_EN : list
        The list of named entities for Ebglish language
    POS : list
        The list of parts of speech
    DEP_RU : list
        The list of dependence labels for Russian language
    DEP_EN : list
        The list of dependence labels for English language
    
    Methods
    -------
    extract
        Applies extraction operation using regexp to every part of the text,
        separated with given delimiter, in every row of the given column
    replace
        Applies replacement using regexp to every part of the text, separated with
        the given delimiter, in every row of the given column
    
    get_uniquetokens
        Extracts a set of unique tokens (words or phrases) from the given column. 
        If attributes 'regexp' and 'repl' are given, it applies them for replacement
        operations to extracted unique tokens joined as an entire text.
    clear_from_label
        Deletes from text data the named entities specified by the parameter 'labels'.
        Named entities can be the certain groups of patterns joined by similar meaning
        (geographic objects, firstnames and lastnames of people, organizations manes).
        A set of labels of named entities is different for different language models
        and is specified in the relevant documentation.
    extract_ents
        Extracts from text data the named entities specified by the parameter 'labels'.
    vect
        Vectorizes text data in the given column and marks the rows which don't have vectors.
    word_extractor
        Extracts the certain number of words from text data in the given column.
        The words must match the specified parameters.
        Use 'full_df'=True to get full dataset with the information about word dependences and parts of the speech.
    mapper
        Maps text data in the given column and the given category using the special dict.
        This dict has bindings between categories and patterns which this method searches in the text data.
        Puts the result in the separate column.
    map_all
        Maps text data in the given column and ALL categories in the special dict.
        Puts the result in the separate columns with the names of categories which are in the dict.
    quoting_stats
        Calculates quoting statistics for the given pattern list.
    nlp_processing
        Applies nlp-processing to text data in the given column.
    """

    def __init__(self, text_col, nlp):
        
        self.nlp = nlp
        self.text_col = text_col
        self.textcol_mod = self.text_col.copy()
        self.LABELS_LIST_RU = ['ORG', 'PER', 'LOC']

        self.LABELS_LIST_EN = ['CARDINAL', 'DATE', 'EVENT', 'FAC', 'GPE', 'LANGUAGE', 'LAW',
                              'LOC', 'MONEY', 'NORP', 'ORDINAL', 'ORG', 'PERCENT', 'PERSON',
                              'PRODUCT', 'QUANTITY', 'TIME', 'WORK_OF_ART']
        
        self.POS = ['NOUN', 'VERB', 'PRON', 'ADJ', 'ADP', 'NUM', 'PUNCT', 'DET', 'ADV',
                    'PROPN', 'AUX', 'CCONJ', 'SCONJ', 'PART', 'X', 'SPACE', 'INTJ']
        
        self.DEP_RU = ['ROOT', 'acl', 'acl:relcl', 'advcl', 'advmod', 'amod', 'appos', 'aux',
                       'aux:pass', 'case', 'cc', 'ccomp', 'compound', 'conj', 'cop', 'csubj',
                       'csubj:pass', 'dep', 'det', 'discourse', 'expl', 'fixed', 'flat',
                       'flat:foreign', 'flat:name', 'iobj', 'list', 'mark', 'nmod', 'nsubj',
                       'nsubj:pass', 'nummod', 'nummod:entity', 'nummod:gov', 'obj', 'obl',
                       'obl:agent', 'orphan', 'parataxis', 'punct', 'xcomp']
        
        self.DEP_EN = ['ROOT', 'acl', 'acomp', 'advcl', 'advmod', 'agent', 'amod', 'appos',
                       'attr', 'aux', 'auxpass', 'case', 'cc', 'ccomp', 'compound', 'conj',
                       'csubj', 'csubjpass', 'dative', 'dep', 'det', 'dobj', 'expl', 'intj',
                       'mark', 'meta', 'neg', 'nmod', 'npadvmod', 'nsubj', 'nsubjpass',
                       'nummod', 'oprd', 'parataxis', 'pcomp', 'pobj', 'poss', 'preconj',
                       'predet', 'prep', 'prt', 'punct', 'quantmod', 'relcl', 'xcomp']
        
    ### ------------------------------------------------------------------------------
    
    ### for text modification in the column
    def replace(self, regexp, text_col=None, update=True, repl='', sep_for_tokens=''):
        """Applies replacement using regexp to every part of the text, separated with given delimiter, in every row of the given column
            
        Parameters
        ----------
        regexp : str
            A regular expression as Python’s raw string notation (r'regexp')
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object for the replacing operation; if None, self.textcol_mod is used (default is None)
        update : bool, optional
            If True, rewrites self.textcol_mod with the result of the replacement (default is True)
        repl : str, optional
            A string for the replacing operation (default is an empty string)
        sep_for_tokens : str, optional
            A delimiter for separating the text on the several parts and then using them for individual processing (replacing operation)
            If '', the whole text is used to process (default is '')
            
        Returns
        -------
        pd.Series
            А column with processed text
        """
        def _replace_(string, regexp, repl, sep_for_tokens):
            if pd.isna(string):
                return string
            if sep_for_tokens:
                lst = str(string).split(sep_for_tokens)
            else:
                lst = [string]
            #
            result = []
            for elem in lst:
                proc_res = re.sub(regexp, repl, elem)
                if proc_res:
                    result.append(proc_res)
            result = sep_for_tokens.join(result)
            return result
 
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod
        
        result = text_col.apply(_replace_, args=(regexp, repl, sep_for_tokens))
        
        if update:
            self.textcol_mod = result
            
        return result
        
    ### ---------------------------------------------------------------------------
    


    ### for text modification in the column
    def extract(self, regexp, text_col=None, update=True, sep_for_tokens=''):
        """Applies extraction operation using regexp to every part of the text, separated with given delimiter, in every row of the given column
            
        Parameters
        ----------
        regexp : str
            A regular expression as Python’s raw string notation (r'regexp')
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object for the extracting operation; if None, self.textcol_mod is used (default is None)
        update : bool, optional
            If True, rewrites self.textcol_mod with the result of the replacement (default is True)
        sep_for_tokens : str, optional
            A delimiter for separating the text on the several parts and then using them for individual processing (replacing operation)
            if '', the whole text is used to process(default is '')
            
        Returns
        -------
        pd.Series
            А column with processed text
        """
        ## func for text extration
        def _extract_(string, regexp, sep_for_tokens=''):
            if pd.isna(string):
                return string
                
            if sep_for_tokens:
                lst = str(string).split(sep_for_tokens)
            else:
                lst = [string]
                
            result = []
            for elem in lst:
                proc_res = ''.join(re.findall(regexp, elem))
                if proc_res:
                    result.append(proc_res)
            result = sep_for_tokens.join(result)
            return result
            
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod
                
        result = text_col.apply(_extract_, args=(regexp, sep_for_tokens))
        
        if update:
            self.textcol_mod = result
        return result

    ### ------------------------------------------------------------------------------
    
    ### getting list of unique patterns from column of texts
    def get_uniquetokens(self, text_col=None, update=True, sep=',', regexp=None, repl=None):

        """ Extracts a set of unique tokens (words or phrases) from the given column. 
            If attributes 'regexp' and 'repl' are given, it applies them for replacement operations to extracted unique tokens joined as an entire text.

        Parameters
        ----------
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.textcol_mod is used (default is None)
        update : bool, optional
            If True, saves the result in self.unique_tokens variable (default is True)
        sep : str, optional
            A delimiter for concatenating the extracted tokens (default is ',')
        regexp : str, optional
            A regular expression as Python’s raw string notation (r'regexp') (default is None)
        repl : str, optional
            A string for the replacing operation (default is an empty string)
                    
        Returns
        -------
        pd.Series
            А column with unique patterns (tokens)
        """

        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod
            
        result = text_col.str.cat(sep=sep)
        
        if regexp and repl:
            for exp, rpl in zip(regexp, repl):
                print(exp, rpl)
                len_before = len(result)
                result = re.sub(exp, rpl, result)
                len_after = len(result)
                print(f'Before: {len_before} ==> After: {len_after}')
                
        result = pd.Series(result.split(sep)).unique()
        result = pd.Series(result)
        result = result[(result!='') & (result!=' ')]
               
        print(f'Final list length: {len(result)}')
        if update:
            self.unique_tokens = result
                
        return result

    ### ------------------------------------------------------------------------------------------
    
    ### for cleaning text in the column from named entity
    def clear_from_label(self, text_col=None, labels='all', update=True, remove='all', filtered=False, aliquot=10):
        """ Deletes from text data the named entities specified by the parameter 'labels'.
            Named entities can be the certain groups of patterns joined by similar meaning
            (geographic objects, firstnames and lastnames of people, organizations manes).
            A set of labels of named entities is different for different language models
            and is specified in the relevant documentation.

        Parameters
        ----------
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.unique_tokens is used (default is None)
        labels : list or 'all'
            If 'all', cleans the text data in the given colum from all named entities,
            otherwise uses the given labels of the named entities (default is 'all')
        update : bool, optional
            If True, rewrites self.unique_tokens with the result of the processing (default is True)
        remove : str, optional
            If 'all', deletes the entire pattern(token) from Series even if one named entity found,
            else if 'every', deletes only the named entities found from a pattern,
            otherwise, deletes the entire pattern from Series,
            if all words in that patterns are specified named entities found(default is 'all')
        filtered : bool, optional
            If True, returns only the filtered result without the statistic information(default is False)
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 10).
                    
        Returns
        -------
        pd.DataFrame
            A dataframe with 2 columns: 'result', 'stats'. 
            'result' contains transformed data (remove = 'every') or original data (remove = 'all' or any other string)
            'stats' contains the special labels obtained as a result of processing. 
            '-1' means - this entire pattern will be deleted from result if 'update' or 'filtered' is True.
            It depends from the parameter 'remove'.
            '0' means that this pattern will remain unchanged
            any other positive number appears only then remove='every' and shows how many words were deleted from the pattern.
            
            Use the parameters 'update' and/or 'filtered' for obtaining the filtered result
        """
        ## built-in func for processing Series
        def _clear_from_label(string, labels, remove, aliquot):

            def check_cond(value_for_check, labels):
                if labels == 'all':
                    return value_for_check.ent_type_ != ''
                                
                return value_for_check.ent_type_ in labels

            # for progress visualization
            if aliquot:
                self.counter += 1
            
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                
            # checking for NaN
            if pd.isna(string):
                return pd.Series([pd.NA, pd.NA])
                
            doc = self.nlp(string)
            stats = 0
            # excluding all text if it has even one token as named entity           
            if remove == 'all':
                if any([check_cond(token, labels) for token in doc]):
                    stats = -1
                result = string

            # filtering text from chosen named entity
            elif remove == 'every':
                res_list = [token.text for token in doc if not check_cond(token, labels) and not token.is_punct]
                result = (' '.join(res_list))
                if result == '':
                    stats = -1
                else:
                    stats = len(doc) - len(res_list)
                    
            # excluding all text if it has all of tokens as named entity                               
            else:
                if all([check_cond(token, labels) for token in doc]):
                    stats = -1
                result = string
            
            return pd.Series([result, stats])

        # checking incoming var and switch to internal varible
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.unique_tokens
            

        # initialization of progress counter
        if aliquot:
            self.iter_lenth = len(text_col)
            self.counter = 0
            self.progr = IntProgress(description='Progress: ', min=0, max=self.iter_lenth)
            self.label = Label(value="0")
            display(widgets.HBox([self.progr, self.label]))

        # processing text in the columns
        cleared_textcol = text_col.apply(_clear_from_label, args=(labels, remove, aliquot))
        cleared_textcol.columns = ['result', 'stats']

        # for update internal var
        if update:
            l_before = len(self.unique_tokens)
            self.unique_tokens = cleared_textcol.loc[cleared_textcol['stats'] != -1, 'result'].reset_index(drop=True)
            l_after = len(self.unique_tokens)
            print(f'Unique tokens: {l_before} => {l_after}')

        # for returning filtered result
        if filtered:
            return cleared_textcol.loc[cleared_textcol['stats'] != -1, 'result'].reset_index(drop=True)

        return cleared_textcol
        
    ### -------------------------------------------------------------------------------------   
    
    ### for extraction named entity from the text in the column
    def extract_ents(self, text_col=None, labels='ru', aliquot=10, sep=',', filtered=False, rest=True, inverse=False):
        """ Extracts from text data the named entities specified by the parameter 'labels'.
            Named entities can be the certain groups of patterns joined by similar meaning
            (geographic objects, firstnames and lastnames of people, organizations manes).
            A set of labels of named entities is different for different language models
            and is specified in the relevant documentation.

        Parameters
        ----------
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.unique_tokens is used (default is None)
        labels : list or 'ru'|'en'
            If 'ru'|'en', extract all defined for the certain language named entities from the text data in the given column,
            otherwise uses the given labels of the named entities (default is 'ru')
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is None).
        sep : str, optional
            A delimiter for result words (default is ',')
        filtered : bool, optional
            If True, returns only filtered series with extracted named entities(default is False).
            You can return a rest text without named entities if you set the parameter 'inverse' to True
        rest : bool, optional
            If True, adds the column with the remained text, after named entities extraction, to the final result (default is True)
        inverse : bool, optional
            If True, uses only with the parameter 'filtered'=True to obtain cleaned from named entities series(default is False)
                            
        Returns
        -------
        pd.DataFrame
            A dataframe with the number of columns equal to the number of named entities found.
            The result shows the extracted named entities sorted by the specified columns.
            If 'rest'=True, adds the column with the remained text after named entities extraction.
            Use the parameters 'filtered' for obtaining the filtered series
        """
        ## built-in func for the processing Series 
        def _extract_ents_w_label(string, labels, aliquot, sep, inverse=False):
                
            # for progress visualization
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                
            # checking for NaN
            if pd.isna(string):
                return np.nan
                
            doc = self.nlp(string)

            if inverse:
                result = [token.text for token in doc if not token.ent_type_ in labels]
            else:
                result = [token.text for token in doc if token.ent_type_ in labels]
            result = (sep.join(result))
            if result == '':
                return np.nan
                
            return result
            
        # defining labels list
        if labels == 'ru':
            labels_list = self.LABELS_LIST_RU
            
        elif labels == 'en':
            labels_list = self.LABELS_LIST_EN
        else:
            labels_list = labels

        # checking incoming var and switch to internal varible
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod

        # for returning concatenated result
        if filtered:
            if aliquot:
                self.iter_lenth = len(text_col)
                self.counter = 0
                self.progr = IntProgress(description='Progress', min=0, max=self.iter_lenth)
                self.label = Label(value="0")
                display(widgets.HBox([self.progr, self.label]))
            return text_col.apply(_extract_ents_w_label, args=(labels_list, aliquot, sep, inverse))
            
        # searching entity with chosen labels
        result = pd.DataFrame()
        for label in self._progress_visual(labels_list, message='Total progr:'):
            
            # initialization of progress counter
            if aliquot:
                self.iter_lenth = len(text_col)
                self.counter = 0
                self.progr = IntProgress(description=f'{label}', min=0, max=self.iter_lenth)
                self.label = Label(value="0")
                display(widgets.HBox([self.progr, self.label]))
            found_ents = text_col.apply(_extract_ents_w_label, args=([label], aliquot, sep))
            print(f'{label} processed')
            found_ents.name = label
            result = pd.concat([result, found_ents], axis=1)

        if rest:
            if aliquot:
                self.iter_lenth = len(text_col)
                self.counter = 0
                self.progr = IntProgress(description='the rest', min=0, max=self.iter_lenth)
                self.label = Label(value="0")
                display(widgets.HBox([self.progr, self.label]))
            found_ents = text_col.apply(_extract_ents_w_label, args=(labels_list, aliquot, sep), inverse=True)
            print('the rest processed')
            found_ents.name = 'rest'
            result = pd.concat([result, found_ents], axis=1)
        
        return result

    ### -------------------------------------------------------------------------------------   
    
    ### for extraction named entity from the text in the column
    def extract_cats(self, text_col=None, labels='all', aliquot=10, df=False, rnd=3):
        """ Extracts category name for from text data.
            Model has to be learnt to find this categories.
           
        Parameters
        ----------
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.unique_tokens is used (default is None)
        labels : list or str
            If 'all', extract all categories from the text data,
            if it is a list of categories, try to find all categories from the list,
            otherwise uses the given labels of the category (default is 'all')
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is None).
        df : str, optional
            Outputs the results as dataframe (default is False)
                                    
        Returns
        -------
        pd.Series or pd.DataFrame
            
            A dataframe with the number of columns equal to the number of categories found.            
        """
        ## built-in func for the processing Series 
        def _extract_cats_w_label(string, labels='all', aliquot=10, rnd=3):
                
            # for progress visualization
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                
            # checking for NaN
            if pd.isna(string):
                return np.nan
                
            doc = self.nlp(string)

            result = doc.cats
            result = {key: round(result[key], rnd) for key in result}
            
            if labels == 'all':
                return result
                
            else:
                if isinstance(labels, (list, tuple, pd.Series)):
                    return {k: v for k, v in result.items() if k in labels}
                elif isinstance(labels, str):
                    if labels in result:
                        return result[labels]
                    else:
                        return np.nan
                else:
                    print('!!!Wrong format for parameter "labels"')
                    print('Use list of str or str')
                    sys.exit()
                    
        # checking incoming var and switch to internal varible
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod

        # for returning concatenated result
        if aliquot:
            self.iter_lenth = len(text_col)
            self.counter = 0
            self.progr = IntProgress(description='Progress', min=0, max=self.iter_lenth)
            self.label = Label(value="0")
            display(widgets.HBox([self.progr, self.label]))
        extr_result = text_col.apply(_extract_cats_w_label, args=(labels, aliquot, rnd))

        if df:
            if isinstance(labels, (list, tuple, pd.Series)) or labels == 'all':
                result = pd.DataFrame.from_records(extr_result)
                result.insert(0, 'text_col', text_col)
            else:
                result = pd.DataFrame({'text_col': text_col, labels: extr_result})
            return result
        return extr_result
        
    ### -----------------------------------------------------------------------------    
    
    def vect(self, text_col=None, nlp=None, update=True, filtered=True, aliquot=10, full_df=False):
        """ Vectorizes text data in the given column and marks the rows which don't have vectors.
            Use 'full_df'=True to obtain the full dataset as result. 
        
        Parameters
        ----------
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.unique_tokens is used (default is None)
        nlp : spacy model class
            If given, it is used to process,
            otherwise self.nlp is used (default is None)
        update : bool, optional
            If True, rewrites self.unique_tokens with the result of the processing 
            and saves the corresponding vectors into self.utokens_vectors(default is True)
        filtered : bool, optional
            If True, returns only filtered result as series where the text data are an index and the corresponding vectors are values.
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 10).
        full_df : bool, optional
            If True, returns the full dataset as result. There are 3 columns in that dataset: text, vectors and flags of availability vectors
                    
        Returns
        -------
        pd.DataFrame or pd.Series

            If 'full_df'=True, returns the full dataset as result.
            There are 3 columns in that dataset: 'text_col', 'vectors', 'has_vectors'.
            If 'full_df'=False, returns series where the text data are an index and the corresponding vectors are values.
        """

        def _vect(string, aliquot):
            
            # for progress visualization
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                
            # checking for NaN
            if pd.isna(string):
                return pd.Series([pd.NA, pd.NA, 0])

            nlp_string = nlp(string)
            if nlp_string.has_vector:
                return pd.Series([string, nlp_string.vector, 1])
                
            return pd.Series([string, nlp_string.vector, 0])
        
        # checking incoming var and switch to internal variable
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.unique_tokens

        if nlp:
            pass
        else:
            nlp = self.nlp

        # for visualization initialization
        if aliquot:
            self.iter_lenth = len(text_col)
            self.counter = 0
            self.progr = IntProgress(description='Progress:', min=0, max=self.iter_lenth)
            self.label = Label(value="0")
            display(widgets.HBox([self.progr, self.label]))

        # text_col prrocessing
        vect_df = text_col.apply(_vect, args=(aliquot,))
        vect_df.columns = ['text_col', 'vectors', 'has_vectors']

        if full_df:
            return vect_df

        if filtered:
            
            filt_textcol = vect_df[vect_df['has_vectors'] == 1]['text_col']
            filt_vect = vect_df[vect_df['has_vectors'] == 1]['vectors']
        else:
            filt_textcol = vect_df['text_col']
            filt_vect = vect_df['vectors']

        if update:
            l_before = len(self.unique_tokens)
            self.unique_tokens = filt_textcol
            self.utokens_vectors = filt_vect
            l_after = len(self.unique_tokens)
            print(f'Unique tokens: {l_before} => {l_after}')
            
        filt_vect.index = filt_textcol
        
        return filt_vect
        
    ### --------------------------------------------------------------------------------------------
    
    def word_extractor(self, pattern=None, text_col=None, threshold=None, count_thres=None,
                       dep=None, pos=None, desc_sim=False, stat=False, full_df=False, aliquot=10):
        """ Extracts the certain number of words from text data in the given column. The words must match the specified parameters.
            Use 'full_df'=True to get full dataset with the information about word dependences and parts of the speech.
                    
        Parameters
        ----------
        pattern : str, optional
            A set of words and phrases which extracted word should be similar to.
            The set is used to calculate similarity metric which can be used to filter and sort words in the text(default=None)
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.textcol_mod is used (default is None)
        threshold : float, optional
            {0, ..., 1} If given, it remains the words, similarity of which is greater than or equal to this threshold(default is None).
        count_thres : int, optional
            Only positive numbers. If given, remains the corresponded quantity of the extracted words.
            If 'pattern' is given, sorts the result by descending similarity(default is None).
        dep : list, optional
            A list of special labels which are sentence dependence labels. Use self.DEP_RU and self.DEP_EN to see them.
            If given, used to filter only the words with these given labels(default is None).
        pos : list, optional
            A list of special labels which are parts of speech. Use self.POS to see them.
            If given, used to filter only the words with these given labels(default is None).
        desc_sim : bool, optional
            If True, changes order of words in the result according to descending similarity(default is False)
        stat : bool, optional
            If True, adds a special column with statistics data for all words as a dict (parts of speech, sentance dependences, similarity)
            (default is False)
        full_df : bool, optional
            If True, returns the full dataset as result.
            There are as many columns as part of speech (pos) and sentence dependence (dep) found there.
            The words corresponding the specified labels are as values in the corresponding columns(default is False).
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 10).
                            
        Returns
        -------
        pd.DataFrame or pd.Series

            If 'full_df'=True, returns the full dataset as result.
            There are columns named as part of speech (pos) and sentence dependence (dep) found there.
            If 'full_df'=False, returns series with extracted words. 
            If 'stat'=True, adds column with statistics and the series turns into dataframe.
        """

        def _word_extractor(string, pattern=None, threshold=None, count_thres=None, dep=None, pos=None,
                            desc_sim=False, stat=False, full_df=False, aliquot=10):
            # for progress visualization
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                
            # checking for NaN
            if pd.isna(string):
                if full_df:
                    return pd.Series({'NOUN': np.nan})
                if stat:
                    return pd.Series([None] * 2)
                return None

            
            #if not nlp_string.has_vector:
            #    if stat:
            #        return pd.Series([None, 'NO VECTORS'])    
            #    return None
                
            if pattern:
                sim_df = pd.DataFrame([(token.text, token.similarity(pattern),
                                        token.dep_, token.pos_) for token in string if token.has_vector],
                                      columns=['tokens', 'similarity', 'dependences', 'poses'])
                if full_df:
                    res_df = pd.concat([sim_df.pivot_table(columns='dependences', values='tokens', aggfunc=','.join),
                                        sim_df.pivot_table(columns='poses', values='tokens', aggfunc=','.join)], axis=1)
                                        
                    res_df = pd.Series(res_df.loc['tokens', :], index=res_df.columns)
                    return res_df
                
                if threshold:
                    result = sim_df[sim_df['similarity'] >= threshold]

            
            else:
                sim_df = pd.DataFrame([(token.text, token.dep_,
                                        token.pos_) for token in string if token.has_vector], columns=['tokens',
                                                                                                           'dependences', 'poses'])
                if full_df:
                    res_df = pd.concat([sim_df.pivot_table(columns='dependences', values='tokens', aggfunc=','.join),
                                        sim_df.pivot_table(columns='poses', values='tokens', aggfunc=','.join)], axis=1)
                                        
                    res_df = pd.Series(res_df.loc['tokens', :], index=res_df.columns)
                    return res_df
                    
                
            result = sim_df
            
                            
            if dep:
                result = result[result['dependences'].isin(dep)]
            
            if pos:
                result = result[result['poses'].isin(pos)]
               
            if count_thres:
                if pattern:
                    sorted_res = result.sort_values('similarity', ascending=False).iloc[:count_thres, :]['tokens']
                    result = result[result['tokens'].isin(sorted_res)]
                else:
                    sorted_res = result.iloc[:count_thres, :]['tokens']
                
                result = result.drop_duplicates(subset='tokens')

            if desc_sim:
                if pattern:
                    result = result.sort_values('similarity', ascending=False)

            result = result['tokens'].str.cat(sep=' ')
            if result == '':
                result = None
            if stat:
                statistics = str(sim_df.to_dict(orient='index'))
                return pd.Series([result, statistics])

            return result

        if pattern:
            nlp_pattern = self.nlp(pattern.lower())
        else:
            nlp_pattern = None

        
        # checking incoming var and switch to internal varible
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
            nlp_textcol = self.nlp_processing(text_col)
            self.word_extractor_nlp = text_col            

        elif isinstance(text_col, spacy.tokens.doc.Doc):
            self.word_extractor_nlp = text_col
            nlp_textcol = text_col
                
        elif hasattr(self, 'word_extractor_nlp'):
            nlp_textcol = self.word_extractor_nlp
                    
        else:
            nlp_textcol = self.nlp_processing(self.textcol_mod)
            self.word_extractor_nlp = nlp_textcol
        
        # for visualization initialization
        if aliquot:
            self.iter_lenth = len(nlp_textcol)
            self.counter = 0
            self.progr = IntProgress(description='Progress:', min=0, max=self.iter_lenth)
            self.label = Label(value="0")
            display(widgets.HBox([self.progr, self.label]))

        # text_col prrocessing
        if stat:
            result = nlp_textcol.apply(_word_extractor, args=(nlp_pattern, threshold, count_thres, dep, pos, desc_sim, stat, full_df, aliquot))
            result.columns = ['sf_result', 'statistics']
            return result

        result = nlp_textcol.apply(_word_extractor, args=(nlp_pattern, threshold, count_thres, dep, pos, desc_sim, stat, full_df, aliquot))
        
        return result

    ### --------------------------------------------------------------------------------------------
    def mapper(self, cat, dict_df, text_col=None, mode='binary', aliquot=1):
        """ Maps text data in the given column and the given category using the special dict.
            This dict has bindings between categories and patterns which this method searches in the text data.
            Puts the result in the separate column.
                    
        Parameters
        ----------
        cat : str
            A category name.
        dict_df : pd.DataFrame
            Dataframe with the dict binding categories and patterns.
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.textcol_mod is used (default is None)
        mode : str, optional
            {binary|patterns|quantity} If 'binary', puts bool values into the result column.
            If 'patterns', puts specified patterns in the result column.
            If 'quantity', puts quantity of the found words into the result column(default is 'binary').
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 10).
                            
        Returns
        -------
        pd.Series

            Series with the values depending on the parameter 'mode' and the same name as the category.
        """
        # defining text column for processing
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod
            
        # filtering patterns for target category
        try:
            pattern_list = dict_df[dict_df['categories'] == cat]['patterns']
        except:
            print('No such categories in the dict')
            exit()

        # matching patterns and text in the column
        result = pd.Series([0] * len(text_col))
        base = pd.Series(['nan'] * len(text_col))
        p_result = base.copy()

        if aliquot:
            patlist_iter = self._progress_visual(pattern_list, aliquot=1)
        else:
            patlist_iter = pattern_list
        for pattern in patlist_iter:
            search_res = text_col.str.contains(rf'\b{re.escape(pattern)}\b', case=False, na=False)
            # returns list of matched patterns for every row
            if mode == 'patterns':
                p_result = p_result.str.cat(base.mask(search_res, pattern), sep=',')
            # returns number of matched patterns for every row
            else:    
                result = result + search_res
                
        if mode == 'patterns':
            result = p_result.str.replace(r'nan,?', '', regex=True)
            result = result.str.replace(r',\Z', '', regex=True)
        # returns binary feature for every row
        elif mode == 'binary':
            result = result > 0
                        
        result.name = cat    
        return result
        
    ### ----------------------------------------------------------------------------------------
    
    def map_all(self, dict_df, text_col=None, mode='quantity', aliquot=1):
        """ Maps text data in the given column and all categories in the special dict.
            This dict has bindings between categories and patterns which this method searches in the text data.
            Puts the result in the separate columns with the names of categories which are in the dict.
                    
        Parameters
        ----------
        dict_df : pd.DataFrame
            Dataframe with the dict binding categories and patterns.
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.textcol_mod is used (default is None)
        mode : str, optional
            {binary|patterns|quantity} If 'binary', puts bool values into the result column.
            If 'patterns', puts specified patterns in the result column.
            If 'quantity', puts quantity of the found words into the result column(default is 'binary').
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 1).
                            
        Returns
        -------
        pd.DataFrame

            DataFrame with the columns corresponding the names of categories and containing the values
            which depending on the parameter 'mode'.
        """
        
        result_df = pd.DataFrame()
        cat_list = list(dict_df['categories'].unique())
        try:
            cat_list.remove('UNKNOWN')
        except:
            pass
        
        if aliquot:
            catlist_iter = self._progress_visual(cat_list, aliquot=1)
        else:
            catlist_iter = cat_list
            
        for cat in catlist_iter:
            result = self.mapper(cat, dict_df, text_col, mode, aliquot=False)
            result_df = pd.concat([result_df, result], axis=1)
        return result_df
        
    ### -----------------------------------------------------------------------------------

    def quoting_stats(self, pattern_list=None, text_col=None, ratio=False):
        """ Calculates quoting statistics for the given pattern list.
                    
        Parameters
        ----------
        pattern_list : list, tuple, pd.Series or None, optional
            A list of the searched patterns, if not given, self.unique_patterns is used (default is None)
        text_col : list, tuple, pd.Series or None, optional
            If given, it is used as an object of the operation; if None, self.textcol_mod is used (default is None)
        ratio : bool, optional
            If True, outputs the result as ratio, otherwise, as a number(default is False).
                                        
        Returns
        -------
        pd.Series

            Series with the length equal to the length of the pattern list.
        """
        # checking incoming var and switch to internal variable
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod
        
        if isinstance(pattern_list, (list, tuple, pd.Series)):
            if isinstance(pattern_list, (list, tuple)):
                pattern_list = pd.Series(pattern_list)
        else:
            pattern_list = self.unique_tokens

        result = []
        for pattern in pattern_list:
            if ratio:
                quotes = t_col.str.contains(rf'\b{re.escape(pattern)}\b').mean()
            else:
                quotes = t_col.str.contains(rf'\b{re.escape(pattern)}\b').sum()
            result.append(quotes)
            
        result = pd.DataFrame({'Patterns': pattern_list, 'Number of quotes': result}).sort_values('Number of quotes', ascending=False)
        return result
        
    ### -----------------------------------------------------------------------------------
    
    def _progress_visual(self, iter_obj, iter_lenth=None, aliquot=1, message='Progress:'):

        if not iter_lenth:
            iter_lenth = len(iter_obj)
        progr =  IntProgress(description=message, min=0, max=iter_lenth)
        label = Label(value="0")
        display(widgets.HBox([progr, label]))
        for i, elem in enumerate(iter_obj, 1):
            yield elem
            if i == 1 or i % aliquot == 0 or i == iter_lenth:
                progr.value = i
                label.value = f'{i} of {str(iter_lenth)}'

    ### -----------------------------------------------------------------------------------

    def nlp_processing(self, text_col, lower=True, aliquot=10):
        """ Applies nlp-processing to text data in the given column.
                    
        Parameters
        ----------
        text_col : list, tuple, pd.Series
            A list of text data used as an object for the processing.
        lower : bool, optional
            If True, converts strings to lowercase before processing(default is True).
        aliquot : int or None, optional
            If specified, showes a progress indicator with updating every specified number of times (aliquot the number),
            if None, doesn't show an indicator(default is 10).
                                               
        Returns
        -------
        pd.Series

            Series with processed text data.
        """
        def _nlp_processing(string, aliquot=10):
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
            if pd.isna(string):
                return np.nan
            if lower:
                return self.nlp(str(string).lower())
            return self.nlp(str(string))
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
             # for visualization initialization
            if aliquot:
                self.iter_lenth = len(text_col)
                self.counter = 0
                self.progr = IntProgress(description='NLP-progress:', min=0, max=self.iter_lenth)
                self.label = Label(value="0")
                display(widgets.HBox([self.progr, self.label]))
            result = text_col.apply(_nlp_processing, args=(aliquot,))
        else:
            result = None
                                
        return result

    ### ----------------------------------------------------------------------------------------

    def get_train_data(self, pattern_list=None, label=None, label_data=None, patterns_convert='ORTH',
                       text_col=None, filtered=True, split=0.1, to_disk='./corpus/', aliquot=10, stratify=None):

        # defining incoming text col
        if isinstance(text_col, (list, tuple, pd.Series)):
            if isinstance(text_col, (list, tuple)):
                text_col = pd.Series(text_col)
        else:
            text_col = self.textcol_mod

        # checking incoming patterns list
        if isinstance(pattern_list, (list, tuple, pd.Series)):
            if patterns_convert:
                print('Transforming to the spacy.matcher format')
                # transform to the spacy.matcher format 
                pattern_list = [[{patterns_convert: pat}] for pat in pattern_list]
            print('Using preconverted patterns list')
        
        elif isinstance(label_data, (list, tuple, pd.Series)):
            print('Using label_data (list or Series with special dict')

        else:
            print('No patterns list or wrong format! No label_data! Nothing to process!\n \
            Use patterns list or label_data in pd.Series or list format')
            
        # func for creating docs with the target labels
        def _labeler(string, label):
            # progress visialization
            if aliquot:
                self.counter += 1
                if self.counter == 1 or self.counter % aliquot == 0 or self.counter == self.iter_lenth:
                    self.progr.value = self.counter
                    self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'
                    
            if pd.isna(string):
                return np.nan
            if string == '':
                return np.nan
            doc = self.nlp(string)
            matches = matcher(doc)
            ents = []
            for match in matches:
                ents.append(Span(doc, match[1], match[2], label=label))
            doc.ents = ents
            return doc

        # func for converting into Docbin format and saving the result to disk
        def _db_transform(nlp_textcol, path):
            db = DocBin()
            for doc in nlp_textcol:
                if pd.notna(doc):
                    db.add(doc)
            db.to_disk(path)
            
        
        
        # annotation text_data using Matcher and doc.ents (for Named Entity Recognition)
        if isinstance(pattern_list, (list, tuple, pd.Series)) and label:
            # for visualization initialization
            if aliquot:
                self.iter_lenth = len(text_col)
                self.counter = 0
                self.progr = IntProgress(description='Progress:', min=0, max=self.iter_lenth)
                self.label = Label(value="0")
                display(widgets.HBox([self.progr, self.label]))
            matcher = Matcher(self.nlp.vocab)
            matcher.add(label, pattern_list)
            print('Matcher initialized successfully')
            result = text_col.apply(_labeler, args=(label,))

        # annotation text_data using doc.cats (for Text Categorization Multilabel)
        elif isinstance(label_data, (list, tuple, pd.Series)):
            result = []
            for text, dict in self._progress_visual(zip(text_col, label_data), iter_lenth=len(text_col), aliquot=aliquot):
                if pd.isna(text):
                    result.append(text)
                else:
                    nlp_text = self.nlp(text)
                    nlp_text.cats = dict
                    result.append(nlp_text)
            result = pd.Series(result)
                
        # filtering n/a
        if filtered:
            result = result[result.notna()]
            
        # spliting train data on train and test(valid) and saving to disk
        if split:
            print(f'Splitting data: TRAIN - {100 - split * 100}%,  TEST - {split * 100}%')
            train, test = train_test_split(result, test_size=split, random_state=1982, shuffle=True, stratify=stratify)
            if to_disk:
                _db_transform(train, f'{to_disk}train.spacy')
                print('Training data locates:')
                print( f'{to_disk}train.spacy')
                _db_transform(test, f'{to_disk}dev.spacy')
                print( f'{to_disk}dev.spacy')
                
            return train, test
            
        return result

class Categorizator:
    """ Provides a set of tools for bulding the dict of bindings between categories and patterns.
        Category is a word, a set of words or phrases that combines the meaning of a certain group of words or texts.
        Pattern is a word or a phrase that is contained in raw text and express a certain meaning.
                    
    Parameters
    ----------
    pattern_list : list, tuple, pd.Series, optional
        A list of patterns. If not given, can be given in the method's parameters.
    cat_list : list, tuple, pd.Series, optional
        A list of categories. If not given, can be given in the method's parameters.
    nlp : spacy model class
        Model for the specified language used for processing. Needed for most of methods.
    quoting : list, tuple, pd.Series or None, optional
        If given, it is used as an object of the quoting calculation; if None, self.textcol_mod is used (default is None)
    only_w_vector : bool, optional
        If True, remains in the result only patterns which have vectors(default is True).
    
    Attributes
    ----------
    nlp : spacy model class
        Model for the specified language using for process.
    pattern_list : pd.Series
        Shows the saved list of patterns.
    cat_list : pd.Series
        Shows the saved list of categories.
    quoting_data : pd.Series
        Shows the saved result of quoting process where patterns are in index
        and the corresponded number of quotes as values.
        
    Methods
    -------
    cat_sim
        Selects patterns from the pattern list that have higher similarity
        than the given threshold for the given category.
    catsim_all
        Applies 'cat_sim' for every category from the categories list.
    pattern_sim
        Selects categories from the categories list that have higher similarity
        than the given threshold for the given pattern.
    patternsim_all
        Applies 'pattern_sim' for every pattern from the pattern list.
    get_quoting
        Calculates quoting statistics for the given pattern list and text column.
    """

    ## func for nlp-processing pd.series
    def textlist_nlp(self, text_list, textlist_name, only_w_vector=True):
       
        def _nlp_proc_(string, only_w_vector=True):

            # for progress visualization
            self.progr.value += 1
            self.label.value = f'{str(self.progr.value)} of {str(self.iter_lenth)}'

            # for checking NA

            if isinstance(string, str):
                # processing func
                nlp_string = self.nlp(string.lower())
                if not only_w_vector:
                    return pd.Series([string, nlp_string])
                if nlp_string.has_vector:
                    return pd.Series([string, nlp_string])
                return pd.Series([None, None])
            else:
                return pd.Series([None, None])

        if isinstance(text_list, (list, tuple, pd.Series)):
            
            # transform into series
            if isinstance(text_list, (list, tuple)):
                text_list = pd.Series(text_list)
            
            # processing patterns
            print(f'Starting NLP-processing for {textlist_name}')
            
            # initialization of progress counter
            self.iter_lenth = len(text_list)
            self.progr =  IntProgress(description='Progress: ', min=0, max=self.iter_lenth)
            self.label = Label(value="0")
            display(widgets.HBox([self.progr, self.label]))

            # processing text_col
            text_result = text_list.apply(_nlp_proc_, args=(only_w_vector,))
            text_result = text_result.dropna() ###
            text_list = text_result.iloc[:, 0]
            nlptexts = text_result.iloc[:, 1]
            print(f'{textlist_name} processed')
            print()
                        
        elif hasattr(self, textlist_name):
            # using preprocessed during initialization patterns
            print(f'Using preprocessed {textlist_name}')
            text_list = eval(f'self.{textlist_name}')
            nlptexts = eval(f'self.nlp{textlist_name.removesuffix("_list")}s')
        else:
            print(f'No {textlist_name.removesuffix("_list")} given. Use param "{textlist_name}".')
            return pd.Series(), pd.Series()
        
        return text_list, nlptexts

    ### ------------------------------------------------------------------------------------------------------------------
         
    ### nlp-processing and saving categories and patterns
    def __init__(self, nlp, pattern_list=None, cat_list=None, quoting=None, only_w_vector=True):
                
        self.nlp = nlp
        
        # nlp-processing of categories
        self.cat_list, self.nlpcats = self.textlist_nlp(cat_list, 'cat_list', only_w_vector=False)

        print(f'Categories without vectors: {self.cat_list.loc[self.nlpcats.transform(lambda x: not x.has_vector)]}')
               
        # nlp-processing of patterns
        self.pattern_list, self.nlppatterns = self.textlist_nlp(pattern_list, 'pattern_list', only_w_vector=only_w_vector)
    
        # quoting calculation
        if isinstance(quoting, (list, tuple, pd.Series)):
            self.quoting_data = self.get_quoting(quoting, self.pattern_list, ratio='both', df=False)
    ### ------------------------------------------------------------------------------------------------------------------
    ## func for similarity calc
    def sim_calc(self, nlp1, nlp2, metric):
        
        if metric == 'cosine':
            sim = float(cosine_similarity(nlp1.vector.reshape(1, -1), nlp2.vector.reshape(1, -1)))
        elif metric == 'euclide':
            sim = -float(euclidean_distances(nlp1.vector.reshape(1, -1), nlp2.vector.reshape(1, -1)))
        else:
            sim = nlp1.similarity(nlp2)
        return sim

    ### ------------------------------------------------------------------------------------------------------------------
    ## advanced func for similarity calc
    def adv_sim_calc(self, nlp1, nlp2, metric='mean'):
        if not metric:
            metric = 'mean'
        calc_dict = {}
        for token2 in nlp2:
            for token1 in nlp1:
                sim = token1.similarity(token2)
                calc_dict[f'{token2.text}_{token1.text}'] = sim
                
        if metric == 'dict':
            return calc_dict

        if metric == 'mean':
            return np.mean(list(calc_dict.values()))

        re_parse = re.match(r'(\Adict_top)([\d.,]+)', metric)
        if re_parse:
            top = int(re_parse.group(2))
            result = dict(sorted(list(calc_dict.items()), key=lambda x: x[1], reverse=True)[:top])
            return result
        
        re_parse = re.match(r'(\Amean_top)([\d.,]+)(_threshold)([\d.,]+)', metric)
        if re_parse:
            top = int(re_parse.group(2))
            threshold = float(re_parse.group(4))
            result = list(calc_dict.values())
            result = np.extract(result > np.float64(threshold), result)
            result = sorted(result, reverse=True)[:top]
            return result
        
        re_parse = re.match(r'(\Amean_top)([\d.,]+)', metric)
        if re_parse:
            top = int(re_parse.group(2))
            result = sorted(list(calc_dict.values()), reverse=True)[:top]
            return sum(result)/len(result)
        
        
        re_parse = re.match(r'(\Amean_quantile)([\d.,]+)(_threshold)([\d.,]+)', metric)
        if re_parse:
            quantile = float(re_parse.group(2))
            threshold = float(re_parse.group(4))
            result = list(calc_dict.values())
            result = np.extract(np.logical_and(result > np.quantile(result, quantile), result > np.float64(threshold)), result).mean()
            return result
        
        re_parse = re.match(r'(\Amean_quantile)([\d.,]+)', metric)
        if re_parse:
            quantile = float(re_parse.group(2))
            result = list(calc_dict.values())
            result = np.extract(result > np.quantile(result, quantile), result).mean()
            return result
            

    ### ------------------------------------------------------------------------------------------------------------------
    
    ### similarity calculation between a category and a set of patterns
    def cat_sim(self, cat, threshold=None, count_thres=None, pattern_list=None, nlppatterns=None, metric=None,
                need_nlp=True, quoting=None, ratio=False, update=False, only_w_vector=True, sim_func='simple'):
        """ Selects patterns from the pattern list that have higher similarity than the given threshold for the given category.
                                
        Parameters
        ----------
        cat : str
            A category for which the method must select relevant patterns.
        threshold : float, optional
            {0, ..., 1} If given, it remains only the patterns, similarity to category of which
            is greater than or equal to this threshold,
            otherwise, remains all patterns.
            If none of patterns has the similarity greater than or equal to the threshold,
            it is marked as 'UNKNOWN'(default is None).
        count_thres : int, optional
            Only positive numbers.
            If given, remains the given quantity of the patterns descending from the top of similarity(default is None).
        pattern_list : list, tuple, pd.Series or None, optional
            If given, it is used as a pattern list, if None, self.pattern_list is used(default is None).
        nlppatterns : list, tuple, pd.Series or None, optional
            If given, it is used as a nlp-preprocessed pattern list,
            otherwise, it is calculated by the built-in func(default is None).
        metric : str, optional
            If None, uses built-in spacy similarity calculation method(cosine),
            if 'cosine', uses the same method from scikit-learn lib,
            if 'euclide', uses euclide distances calculation method from scikit-learn(default is None),
            if 'dict', returns the dictionary with calculated similarity between all tokens,
            if 'mean', returns the mean of data in the dict above,
            if 'dict_top{number}' where {number} is a number of top elements which will be returned,
            if 'mean_top{number}' where {number} is a number of top elements which are averaging,
            if 'mean_top{number1}_threshold{number2}' - here we specify two condition to filter the results which are averaging then,
            if 'mean_quantile{number}' - here we use quantile to filter the results which are averaging then
            if 'mean_quantile{number1}_threshold{number2}' - here we use quantile and threshold to filter the results which are averaging then.
        need_nlp : bool, optional
            If True, nlp-preprocesses a cat and patterns(default is True).
        quoting : list, tuple, pd.Series or None, optional
            If given, uses for the quoting calculation,if None, self.textcol_mod is used (default is None).
        ratio : bool, optional
            If True, outputs the similarity result as ratio, otherwise, as a number(default is False).
        update : bool, optional
            If True, updates self.pattern_list, self.nlppatterns(default is False).
        only_w_vector : bool, optional
            If True, remains in the result only patterns which have vectors(default is True).
        sim_func : str, optional
            It allows to specify the type of similarity calculation function.
            If 'simple' - self.sim_calc is used,
            if 'advanced' - self.adv_sim_calc is used.
                        
        Returns
        -------
        pd.DataFrame
            A dataframe with 3 columns: 'patterns', 'number of quotes' and a name of the given category.
            The column with the name of the given category contains similarity data.
            'number of quotes' contains quantities of text row in the given target column which match patterns.
            If there is no data for quoting calculation, the dataframe has only 2 columns(without 'number of quotes')
        """
         
        if need_nlp:
            # nlp-processing of a category
            cat_mod = self.nlp(cat.lower())

            if not cat_mod.has_vector:
                print('No vectors for this category')
                return pd.DataFrame({'patterns': ['NO VECTORS'], cat: [0]})
            
            # nlp-processing of patterns
            pattern_list, nlppatterns = self.textlist_nlp(pattern_list, 'pattern_list', only_w_vector)
            
        else:
            cat_mod = cat
            
        if update:
            self.pattern_list, self.nlppatterns = pattern_list, nlppatterns
        
        # similarity calculation
        if sim_func == 'simple':
            sim_list = nlppatterns.apply(self.sim_calc, args=(cat_mod, metric))
            
        else:
            sim_list = nlppatterns.apply(self.adv_sim_calc, args=(cat_mod, metric))

        # result outputs
        if  metric != 'dict':
            result = pd.DataFrame({'patterns': pattern_list, cat: sim_list}).sort_values(cat, ascending=False)
            
        else:
            result = pd.DataFrame({'patterns': pattern_list, cat: sim_list})

        # filtering by similarity threshold
        if threshold:
            result = result[result[cat] >= threshold]

        # cutting by number of elements in result
        if count_thres:
            result = result.iloc[:count_thres, :]

        if need_nlp: 
            # quoting calculation and adding to result (NOT TO USING WITH catsim_all)
            mapping_ser = self.get_quoting(quoting, pattern_list, ratio)
                    
            if not mapping_ser.empty:
                print('Mapping with quoting data...')
                result['number of quotes'] = result['patterns'].map(mapping_ser)
                result = result.sort_values([cat,
                                             'number of quotes',
                                             'patterns'], ascending=False)[['patterns', cat,
                                                                            'number of quotes']]
                if update:
                    self.quoting_data = mapping_ser
        # checking for emptiness
        if result.empty:
            if need_nlp:
                print('Result is empty!')
            result = pd.DataFrame({'patterns': ['UNKNOWN'], cat: [0]})
            
        return result

    ### ------------------------------------------------------------------------------------------------------------------
    
    ### similarity calculation between a set of categories and a set of patterns (sorting by categories)
    def catsim_all(self, cat_list=None, threshold=None, count_thres=None, pattern_list=None,
                   metric=None, quoting=None, ratio=False, update=False, only_w_vector=True):
        """ Selects patterns from the pattern list that have higher similarity than the given threshold
        for EVERY category from categories list.
                                
        Parameters
        ----------
        cat_list : list, tuple, pd.Series or None, optional
            A list of categories for which the method must select relevant patterns.
            If not given, self.cat_list, self.nlpcats are used(default is None)
        threshold : float, optional
            {0, ..., 1} If given, it remains only the patterns, similarity to category of which
            is greater than or equal to this threshold,
            otherwise, remains all patterns.
            If none of patterns has the similarity greater than or equal to the threshold,
            it is marked as 'UNKNOWN'(default is None).
        count_thres : int, optional
            Only positive numbers.
            If given, remains the given quantity of the patterns descending from the top of similarity(default is None).
        pattern_list : list, tuple, pd.Series or None, optional
            If given, it is used as a pattern list, if None, self.pattern_list, self.nlppatterns are used(default is None).
        metric : str, optional
            If None, uses built-in spacy similarity calculation method(cosine),
            if 'cosine', uses the same method from scikit-learn lib,
            if 'euclide', uses euclide distances calculation method from scikit-learn(default is None).
        quoting : list, tuple, pd.Series or None, optional
            If given, uses for the quoting calculation,if None, self.textcol_mod is used (default is None).
        ratio : bool, optional
            If True, outputs the similarity result as ratio, otherwise, as a number(default is False).
        update : bool, optional
            If True, updates self.pattern_list, self.nlppatterns, self.cat_list, self.nlpcats(default is False).
        only_w_vector : bool, optional
            If True, remains in the result only patterns which have vectors(default is True).
                        
        Returns
        -------
        pd.DataFrame
            A dataframe with 4 columns: 'categories', 'patterns', 'similarity' and 'number of quotes'.
            'number of quotes' contains quantities of text row which match patterns.
            If there is no data for quoting calculation, the dataframe has only 3 columns(without 'number of quotes')
        """
        
        result = pd.DataFrame()

        # nlp-processing of categories
        cat_list, nlpcats = self.textlist_nlp(cat_list, 'cat_list', only_w_vector=False)

        # nlp-processing of patterns
        pattern_list, nlppatterns = self.textlist_nlp(pattern_list, 'pattern_list', only_w_vector)

        if update:
            self.cat_list, self.nlpcats = cat_list, nlpcats
            self.pattern_list, self.nlppatterns = pattern_list, nlppatterns

        # similarity calculation and concatenation results
        for cat, nlpcat in self._progress_visual(zip(cat_list, nlpcats),
                                                 iter_lenth=len(cat_list), aliquot=10):
            sim_result = self.cat_sim(nlpcat, threshold, count_thres,
                                      pattern_list, nlppatterns, metric, need_nlp=False)  
            sim_result = sim_result.rename(columns={nlpcat: 'similarity'})
            sim_result['categories'] = cat
            result = pd.concat([result, sim_result])
        print('Successfully.')

        # mapping with quoting data
        mapping_ser = self.get_quoting(quoting, pattern_list, ratio)
        if not mapping_ser.empty:
            result['number of quotes'] = result['patterns'].map(mapping_ser)
            result = result.sort_values(['categories',
                                         'similarity',
                                         'number of quotes'], ascending=[True, False, False])[['categories',
                                                                                               'patterns',
                                                                                               'similarity',
                                                                                               'number of quotes']]

        else:
            result = result.sort_values(['categories',
                                         'similarity'], ascending=[True, False])[['categories',
                                                                                  'patterns',
                                                                                  'similarity']]
        return result

    ### ------------------------------------------------------------------------------------------------------------------
    
    ### similarity calculation between a pattern and a set of categories
    def pattern_sim(self, pattern, threshold=None, count_thres=None, cat_list=None, nlpcats=None, metric=None,
                    need_nlp=True, update=False, only_w_vector=True):
        """ Selects categories from the categories list that have higher similarity than the given threshold for the given pattern.
                                
        Parameters
        ----------
        pattern : str
            A pattern for which the method must select relevant categories.
        threshold : float, optional
            {0, ..., 1} If given, it remains only the categories, similarity to a pattern of which
            is greater than or equal to this threshold,
            otherwise, remains all categories.
            If none of categories has the similarity greater than or equal to the threshold,
            it is marked as 'UNKNOWN'(default is None).
        count_thres : int, optional
            Only positive numbers.
            If given, remains the given quantity of the categories descending from the top of similarity(default is None).
        cat_list : list, tuple, pd.Series or None, optional
            If given, it is used as a categories list, if None, self.cat_list is used(default is None).
        nlpcats : list, tuple, pd.Series or None, optional
            If given, it is used as a nlp-preprocessed categories list,
            otherwise, it is calculated by the built-in func(default is None).
        metric : str, optional
            If None, uses built-in spacy similarity calculation method(cosine),
            if 'cosine', uses the same method from scikit-learn lib,
            if 'euclide', uses euclide distances calculation method from scikit-learn(default is None).
        need_nlp : bool, optional
            If True, nlp-preprocesses a pattern and categories(default is True).
        update : bool, optional
            If True, updates self.cat_list, self.nlpcats(default is False).
        only_w_vector : bool, optional
            If True, remains in the result only categories which have vectors(default is True).
                        
        Returns
        -------
        pd.DataFrame
            A dataframe with 2 columns: 'categories' and a name of the given pattern.
            The column with the name of the given pattern contains similarity data.
        """
        if need_nlp:
            # nlp-processing of patterns
            pattern_mod = self.nlp(pattern.lower())
            if not pattern_mod.has_vector:
                print('No vectors for this pattern')
                return pd.DataFrame({'categories': ['NO VECTORS'], pattern: [0]})

            # nlp-processing of a category
            cat_list, nlpcats = self.textlist_nlp(cat_list, 'cat_list', only_w_vector)
            
            

        else:
            pattern_mod = pattern
         
        if update:
            self.cat_list, self.nlpcats = cat_list, nlpcats
        
        # similarity calculation
        sim_list = nlpcats.apply(self.sim_calc, args=(pattern_mod, metric))

        # result outputs
        result = pd.DataFrame({'categories': cat_list, pattern: sim_list}).sort_values(pattern, ascending=False)

        # filtering by similarity threshold
        if threshold:
            result = result[result[pattern] >= threshold]

        # cutting by number of elements in result
        if count_thres:
            result = result.iloc[:count_thres, :]

        # checking for emptiness
        if result.empty:
            if need_nlp:
                print('Result is empty!')
            result = pd.DataFrame({'categories': ['UNKNOWN'], pattern: [0]})
            
        return result

    ### ------------------------------------------------------------------------------------------------------------------
    
    ### similarity calculation between a set of patterns and a set of categories (sorting by quoting)
    def patternsim_all(self, pattern_list=None, threshold=None, count_thres=None, cat_list=None,
                       metric=None, quoting=None, ratio=False, only_w_vector=True):
        """ Selects categories from the categories list that have higher similarity than the given threshold
        for EVERY pattern from pattern list.
                                
        Parameters
        ----------
        pattern_list : list, tuple, pd.Series or None, optional
            A list of patterns for which the method must select relevant categories.
            If not given, self.pattern_list, self.nlppatterns are used(default is None)
        threshold : float, optional
            {0, ..., 1} If given, it remains only the categories, similarity to a pattern of which
            is greater than or equal to this threshold,
            otherwise, remains all categories.
            If none of categories has the similarity greater than or equal to the threshold,
            it is marked as 'UNKNOWN'(default is None).
        count_thres : int, optional
            Only positive numbers.
            If given, remains the given quantity of the categories descending from the top of similarity(default is None).
        cat_list : list, tuple, pd.Series or None, optional
            If given, it is used as a categories list, if None, self.cat_list, self.nlpcats are used(default is None).
        metric : str, optional
            If None, uses built-in spacy similarity calculation method(cosine),
            if 'cosine', uses the same method from scikit-learn lib,
            if 'euclide', uses euclide distances calculation method from scikit-learn(default is None).
        quoting : list, tuple, pd.Series or None, optional
            If given, uses for the quoting calculation,if None, self.textcol_mod is used (default is None).
        ratio : bool, optional
            If True, outputs the similarity result as ratio, otherwise, as a number(default is False).
        update : bool, optional
            If True, updates self.pattern_list, self.nlppatterns, self.cat_list, self.nlpcats(default is False).
        only_w_vector : bool, optional
            If True, remains in the result only categories which have vectors(default is True).
                        
        Returns
        -------
        pd.DataFrame
            A dataframe with 4 columns: 'patterns', 'categories', 'similarity' and 'number of quotes'.
            'number of quotes' contains quantities of text row which match patterns.
            If there is no data for quoting calculation, the dataframe has only 3 columns(without 'number of quotes')
        """

        result = pd.DataFrame()

        # nlp-processing of categories
        cat_list, nlpcats = self.textlist_nlp(cat_list, 'cat_list', only_w_vector=False)

        # nlp-processing of patterns
        pattern_list, nlppatterns = self.textlist_nlp(pattern_list, 'pattern_list', only_w_vector)

        # similarity calculation and concatenation results
        print('Calculating similarity...')
        for pattern, nlppattern in self._progress_visual(zip(pattern_list, nlppatterns),
                                                         iter_lenth=len(pattern_list), aliquot=10):
            sim_result = self.pattern_sim(nlppattern, threshold, count_thres,
                                          cat_list, nlpcats, metric, need_nlp=False)  
            sim_result = sim_result.rename(columns={nlppattern: 'similarity'})
            sim_result['patterns'] = pattern
            result = pd.concat([result, sim_result])
        print('Successfully.')
        
        # mapping with quoting data
        mapping_ser = self.get_quoting(quoting, pattern_list, ratio)
                
        if not mapping_ser.empty:
            result['number of quotes'] = result['patterns'].map(mapping_ser)
            result = result.sort_values(['number of quotes',
                                         'patterns',
                                         'similarity'], ascending=[False, True, False])[['patterns',
                                                                                         'categories',
                                                                                         'similarity',
                                                                                         'number of quotes']]

        else:
            result = result.sort_values(['patterns',
                                         'similarity'], ascending=[True, False])[['patterns',
                                                                                  'categories',
                                                                                  'similarity']]
        return result

    ### ------------------------------------------------------------------------------------------------------------------
    
    ### quoting calculation
    def get_quoting(self, text_col, pattern_list=None, ratio=False, df=False):
        """ Calculates a number of quotes in the given text column for every pattern in the patterns list.
        
        Parameters
        ----------
        text_col : list, tuple, pd.Series or None
            If given, it is used as an object for the operation; if None, self.textcol_mod is used.
        pattern_list : list, tuple, pd.Series or None, optional
            A list of patterns which the method must find and count in the given text column(default is None).
            If not given, self.pattern_list is used(default is None)
        ratio : bool or str, optional
            If True, outputs the similarity result as ratio, otherwise, as a number.
            If 'both', outputs the ratio and the number in the separate columns(default is False).
        df : bool, optional
            If True, returns the result as pd.DataFrame(default is False).
                                
        Returns
        -------
        pd.Series or pd.DataFrame
            If 'df=True' and 'ratio='both'', result is a DataFrame with 3 columns:
            'patterns', 'number of quotes', 'quotes ratio'.
            If 'df=False' and 'ratio='both'', result is a DataFrame with 2 columns:
            'number of quotes', 'quotes ratio' and patterns as index.
            If 'df=True' and 'ratio<>'both'', result is a DataFrame with 2 columns:
            'patterns' and 'number of quotes' or 'quotes ratio'.
            If 'df=False' and 'ratio<>'both'', result is a Series.
        """
        print('Starting quotes counting...')
        if isinstance(text_col, (pd.Series, list, tuple)):
            # transform into series
            if not isinstance(text_col, (pd.Series)):
                text_col = pd.Series(text_col)
            #mapping_ser = self.get_quoting(quoting, pattern_list, ratio)

            if isinstance(pattern_list, (pd.Series, list, tuple)):
                if isinstance(pattern_list, (list, tuple)):
                    pattern_list = pd.Series(pattern_list)


            elif hasattr(self, 'pattern_list'):
                # using preprocessed during initialization patterns
                print('Using preprocessed pattern_list')
                pattern_list = self.pattern_list
                         
            else:
                return print('No pattern given. Use param "pattern_list".')
        
            
            quotes = []
            quotes_ratio = []
            if pattern_list.empty:
                return print('Pattern list is empty!')
            else:
                for pattern in self._progress_visual(pattern_list, aliquot=10):
                    if ratio == 'both':
                        stats = text_col.str.contains(rf'\b{re.escape(pattern)}\b', case=False).agg(['sum', 'mean'])
                        quotes.append(stats.iloc[0])
                        quotes_ratio.append(stats.iloc[1])
                    elif ratio:
                        stats = text_col.str.contains(rf'\b{re.escape(pattern)}\b', case=False).mean()
                        quotes.append(stats)
                    else:
                        stats = text_col.str.contains(rf'\b{re.escape(pattern)}\b', case=False).sum()
                        quotes.append(stats)
                                                                       
                print('Successfully.')
                #print(quotes)
                
                if df:
                    if ratio == 'both':
                        return pd.DataFrame({'patterns': pattern_list,
                                             'number of quotes': quotes,
                                             'quotes ratio': quotes_ratio}).sort_values('number of quotes', ascending=False)
                    result = pd.DataFrame({'patterns': pattern_list,
                                           'number of quotes': quotes}).sort_values('number of quotes', ascending=False)
                    return result
                    
                if ratio == 'both':
                    return pd.DataFrame({'number of quotes': quotes, 'quotes ratio': quotes_ratio}, index=pattern_list)
                quotes = pd.Series(quotes, index=pattern_list)
                return quotes
        
        elif hasattr(self, 'quoting_data'):
            print('Using preprocessed quoting data')
            if ratio:
                return self.quoting_data['quotes ratio']
            return self.quoting_data['number of quotes']
            
        else:
            print('No text data for quoting! Uze param "quoting".')
            return pd.Series()
  
    ### ------------------------------------------------------------------------------------------------------------------
    
    def _progress_visual(self, iter_obj, iter_lenth=None, aliquot=1, message='Progress:'):

        if not iter_lenth:
            iter_lenth = len(iter_obj)
        progr =  IntProgress(description=message, min=0, max=iter_lenth)
        label = Label(value="0")
        #total_label = Label(value=f' of {str(iter_lenth)}')
        display(widgets.HBox([progr, label]))
        for i, elem in enumerate(iter_obj, 1):
            if i == 1 or i % aliquot == 0 or i == iter_lenth:
                progr.value = i
                label.value = f'{i} of {str(iter_lenth)}'
            yield elem
   