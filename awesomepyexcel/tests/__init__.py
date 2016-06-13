# -*- encoding: utf-8 -*-
import os
from unittest import TestCase

from awesomepyexcel.core import Book, Field

class MyBook(Book):
    headers = [
        Field(u'№', is_counter=True),
        Field(u"Название", key='name'),
        Field(u'Количество букв в названии', key=lambda obj: len(obj['name'])),
        Field(u'Число', key='counter')
    ]

class ExcelTestCase(TestCase):
    def test_main(self):
        data = [
            {'name': u"ABCDE", 'counter': 12},
            {'name': u"FDEG", 'counter': 30},
            {'name': u"FF", 'counter': 3},
            {'name': u"ABAS", 'counter': 4},
            {'name': u"SDF", 'counter': 12},
            {'name': u"FDSDFSEG", 'counter': 33},
        ]
        file = MyBook(data).save()
        self.assertTrue(os.path.isfile(file))
        os.remove(file)
