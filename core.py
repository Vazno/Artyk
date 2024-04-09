# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import os
from collections import Counter
from typing import List

import spacy

from path_utils import resource_path
from download_lemmatizers import models

def generate_co_occurrence_matrix(graph: List[List[str]], binary:bool=False):
    '''Generate co-occurrence matrix based on undirected graph.'''
    all_keywords = list()
    for line in graph:
        for keyword in line:
            all_keywords.append(keyword)
    
    # Sorting by quantity
    c = Counter(all_keywords)
    most_common = c.most_common()
    sorted_arr = list()
    for keyword, frequency in most_common:
        sorted_arr.append(keyword)

    matrix = list()
    matrix.append(list())
    matrix[-1].append(None)

    # Filling first row with var names
    for keyword in sorted_arr:
        matrix[-1].append(keyword)

    # Filling first vertical column with var names.
    # Not filling part of the matrix diagonally
    # so that we don't have to calculate twice since
    # matrix is symmetrical
    symmetry = 0

    for keyword in sorted_arr:
        matrix.append(list())
        matrix[-1].append(keyword)

        for j in range(symmetry):
            matrix[-1].append("") # Leaving empty for auto symmetry fill later

        for i in range(len(sorted_arr)-symmetry):
            matrix[-1].append(0)
        symmetry += 1
    
    # Counting occurence for each element in matrix
    for y in range(len(matrix)):
        for x in range(len(matrix[y])):
            if y != 0 and x != 0:
                if matrix[y][x] == "":
                    # Pass for symmetry fill later
                    pass
                elif x == y: # Since matrix[y][x] is the two same keywords, so it's 0 by default
                    pass
                else:
                    for keywords in graph:
                        if matrix[y][0] in set(keywords) and matrix[x][0] in set(keywords):
                            matrix[y][x] += 1
                            if binary:
                                break

    # Symmetry fill
    for y in range(len(matrix)):
        for x in range(len(matrix[y])):
            if matrix[y][x] == "":
                matrix[y][x] = matrix[x][y]
    return matrix

def homogenize(graph: List[List[str]], lemmatize_: bool=True, language:str="english") -> List[List[str]]:
    '''Homogenize strings for co-occurrence analysis.
    Converts strings in list in list to their lower-cased and lemmatized version'''
    homogenized_words = list(list())
    i = 0

    model = models[language.lower()]
    nlp = spacy.load(os.path.join(resource_path("models"), model))

    if lemmatize_:
        for line in graph:
            i += 1
            homogenized_words.append(list())
            for text in line:
                text = text.lower()
                doc = nlp(text)
                lemmas = [token.lemma_ for token in doc]
                s = " ".join(lemmas)
                homogenized_words[-1].append(s)
    else:
        for line in graph:
            i += 1
            homogenized_words.append(list())
            for text in line[0]:
                homogenized_words[-1].append(text.lower())
    return homogenized_words

def exclude_keywords_from_graph(graph: List[List[str]], exclude_keywords: List[str]) -> List[List[str]]:
    '''Returns graph where given keywords are excluded from graph (Nodes connected to the excluded keywords (nodes) are removed too).'''
    fixed_graph = list()
    if exclude_keywords == None:
        return graph
    lower_cased = [word.lower().strip() for word in exclude_keywords]

    for line in graph:
        if len(set(line).intersection(set(lower_cased))) != 0:
            pass
        else:
            fixed_graph.append(line)
    return fixed_graph
