from collections import Counter
from typing import List, Union

import openpyxl

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
                        if matrix[y][0] in keywords and matrix[x][0] in keywords:
                            matrix[y][x] += 1
                            if binary:
                                break

    # Symmetry fill
    for y in range(len(matrix)):
        for x in range(len(matrix[y])):
            if y != 0 and x != 0:
                if matrix[y][x] == "":
                    matrix[y][x] = matrix[x][y]
    return matrix


def generate_excel(matrix: List[List[Union[str, int]]],
                   output_filename: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    for arr in matrix:
        ws.append(arr)
    wb.save(output_filename)
