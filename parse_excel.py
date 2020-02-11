import os
import xlrd

NAME_COL = 0
VENMO_COL = 1
LOCATION_COL = 2
STICKER_START_COL = 3

class Person:
   def __init__(self, name, venmo, location, stickers):
      self.name = name
      self.venmo = venmo
      self.location = location
      self.stickers = stickers


   def __eq__(self, other):
      return self.name == other.name and self.venmo == other.venmo


   def __str__(self):
      return self.name + "," + self.venmo + "," + self.location


class Sticker:
   def __init__(self, name, quantity, url):
      self.name = name
      self.quantity = quantity
      self.url = url


   def __eq__(self, other):
      return self.name == other.name and self.url == other.url


   def __str__(self):
      return "%s,%d" % (self.name, self.quantity)


   def full_str(self):
      return "%s,%d,%s" % (self.name, self.quantity, self.url)


def get_sticker_name(url):
   raw_name = url[url.find("works/") + 6 : url.find("?")]
   return raw_name[raw_name.find("-") + 1 :]


def update_stickers(new_stickers, stickers):
   for new_sticker in new_stickers:
      if new_sticker in stickers:
         index = stickers.index(new_sticker)
         stickers[index].quantity += new_sticker.quantity
      else:
         stickers.append(new_sticker)


def update_people_stickers(new_person, people):
   if new_person in people:
      index = people.index(new_person)
      update_stickers(new_person.stickers, people[index].stickers)
   else:
      people.append(new_person)


def copy_sticker(sticker):
   return Sticker(sticker.name, sticker.quantity, sticker.url)


def parse_xlsx(file_path):
   people = []
   all_stickers = []
   workbook = xlrd.open_workbook(file_path, ragged_rows = True)
   sheet = workbook.sheet_by_index(0)

   for row in range(sheet.nrows):
      cur_person = Person(sheet.cell_value(row, NAME_COL), sheet.cell_value(row, VENMO_COL), sheet.cell_value(row, LOCATION_COL), [])
      cur_sticker_list = []

      for col in range(STICKER_START_COL, sheet.row_len(row), 2):
         sticker_url = sheet.cell_value(row, col)
         sticker_quantity = int(str(sheet.cell_value(row, col + 1))[:1])
         cur_sticker = Sticker(get_sticker_name(sticker_url), sticker_quantity, sticker_url)

         cur_sticker_list.append(cur_sticker)

      update_stickers(cur_sticker_list, cur_person.stickers)
      update_people_stickers(cur_person, people)
      update_stickers([copy_sticker(sticker) for sticker in cur_sticker_list], all_stickers)

   return people, all_stickers


def get_biggest_order(people):
   biggest = -1
   for person in people:
      biggest = max(biggest, len(person.stickers))
   return biggest


def generate_people_header(people):
   header = "Name,Venmo,Location,"
   sticker_qtys = ["Sticker %d,Qty %d" % (i, i) for i in range(1, get_biggest_order(people) + 1)]
   return header + ",".join(sticker_qtys)


def generate_stickers_header():
   return "Name,Total Qty,URL"


def generate_AGG(filename, people, all_stickers):
   lines = []
   lines.append(generate_people_header(people) + "\n")
   for person in people:
      cur_line = str(person) + ","
      cur_line = cur_line + ",".join([str(person_sticker) for person_sticker in person.stickers])
      lines.append(cur_line + "\n")

   lines.append("\n\n" + generate_stickers_header() + "\n")
   for sticker in all_stickers:
      lines.append(sticker.full_str() + "\n")

   with open(os.getcwd() + "\\" + "AGG_" + filename[: filename.find(".xlsx")] + ".csv", "w") as output:
      output.writelines(lines)
   output.close()



def main():
   for filename in os.listdir(os.getcwd()):
      if filename.endswith(".xlsx") and not filename.startswith("AGG_"):
         workbook_path = os.getcwd() + "\\" + filename
         people, stickers = parse_xlsx(workbook_path)

         generate_AGG(filename, people, stickers)



if __name__ == "__main__":
   main()
