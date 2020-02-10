import unittest
from parse_excel import *

class TestParseExcel(unittest.TestCase):

   def test_get_sticker_name(self):
      self.assertEqual(get_sticker_name("https://www.redbubble.com/people/user/works/11111111-name?p=sticker"), "name")
      self.assertEqual(get_sticker_name("https://www.redbubble.com/people/user/works/22222222-other-name?kind=sticker"), "other-name")


   def test_update_stickers(self):
      old_stickers = [Sticker("a", 1, "aaaaa"), Sticker("b", 2, "bbbbb")]
      new_stickers = [Sticker("a", 3, "aaaaa"), Sticker("b", 3, "ccccc")]
      update_stickers(new_stickers, old_stickers)
      self.assertEqual(old_stickers, [Sticker("a", 4, "aaaaa"), Sticker("b", 2, "bbbbb"), Sticker("b", 3, "ccccc")])


   def test_update_people(self):
      old_people = [Person("Bob", "Bob-B", [Sticker("a", 1, "aaaaa")])]

      new_person_1 = Person("Bob", "Bob-B", [Sticker("a", 1, "aaaaa"), Sticker("b", 3, "bbbbb")])
      update_people_stickers(new_person_1, old_people)
      self.assertEqual(old_people, [Person("Bob", "Bob-B", [Sticker("a", 2, "aaaaa"), Sticker("b", 3, "bbbbb")])])

      new_person_2 = Person("Alice", "Alice-A", [Sticker("a", 1, "aaaaa")])
      update_people_stickers(new_person_2, old_people)
      self.assertEqual(old_people, [Person("Bob", "Bob-B", [Sticker("a", 2, "aaaaa"), Sticker("b", 3, "bbbbb")]), Person("Alice", "Alice-A", [Sticker("a", 1, "aaaaa")])])


if __name__ == '__main__':
   unittest.main()