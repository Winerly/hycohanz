# -*- coding: utf-8 -*-
# author: JiemingWang
"""
Build a class named box and it includes some points that could be used in HFSS scripts.
"""

import numpy as np
import numpy.core.defchararray as npchar

class Box():
    def __init__(self, name, unit, start_point, box_size):

        '''

        :param name: str
            The box name in HFSS.
        :param unit: str
            The unit of x,y and z axes.
        :param start_point: a float list of length=3.
        :param box_size: a float list of length=3.
            These are parameters that you need to create a HFSS box.
        '''

        # Get the left-rear-down point as start point,
        # then the box_size numbers are all positive.

        start_point = np.array(start_point)
        box_size = np.array(box_size)
        box_size_sign = box_size < 0
        start_point = start_point + box_size * box_size_sign
        box_size = abs(box_size)

        # Get the matrix [[0,0,0],[1,0,0],[1,1,0]...[1,1,1]] and obtain all vertexes.

        l = np.broadcast_to(np.arange(8), (3,8))
        mulcator = np.bitwise_and(l.T, np.array([1,2,4]))
        mulcator = np.right_shift(mulcator, np.arange(3))

        vertexes = start_point + mulcator * box_size
        self.name = name
        self.unit = unit
        self.vertexes = vertexes

    def get_face_point(self,direction):

        '''
        Get a coordinate of a point in a face of the box
        :param direction: str
            It indicates which face of box need to calculate,valid inputs:'left','right','up','down','front','rear'.
        :return: a numpy array of strings,length=3
            The coordinate of face center.
        '''

        v_indexes = {'left':(1,4),
                    'right':(3,6),
                    'up':(4,7),
                    'down':(0,3),
                    'front':(1,7),
                    'rear':(0,6)}.get(direction,(-1,-1))

        if v_indexes == (-1,-1):
            raise NameError

        face_point = self.vertexes[v_indexes[0]] + 0.5 * \
                     (self.vertexes[v_indexes[1]] - self.vertexes[v_indexes[0]])
        return npchar.add(face_point.astype('str'), self.unit)

    def get_face_edge(self,direction):

        '''
        Get the up and left edges of a face of the box.
        :param direction: str
            It indicates which face of box need to calculate,valid inputs:'left','right','up','down','front','rear'.
        :return: a numpy array of strings,size=(4,3)
            The coordinates of face left-up,right-up,left-up and left-down vertexes.
        '''

        v_indexes = {'left': np.array([5, 4, 0]),
                     'right': np.array([7, 6, 3]),
                     'up': np.array([4, 6, 5]),
                     'down': np.array([0, 2, 1]),
                     'front': np.array([5, 7, 1]),
                     'rear': np.array([4, 6, 0])}.get(direction, np.array([-1, -1, -1]))

        if v_indexes[0] == -1:
            raise NameError

        edge_point = self.vertexes[v_indexes[[0, 1, 0, 2]]]
        return npchar.add(edge_point.astype('str'), self.unit)


