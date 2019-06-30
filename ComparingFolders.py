from os import getcwd, path
from sys import argv, exit
import win32com.client
import os

def die(message):
    print (message)
    exit(1)

def cmp(original_folder, modified_folder):
  print('Working...')
  wordapp = win32com.client.Dispatch("Word.Application")  # Create new Word Object
  wordapp.Visible = 0  # Word Application should`t be visible
  for dirpath, dirname, filename in os.walk(original_folder):
      original_list = list(filename)
  for dirpath2, dirname2, filename2 in os.walk(modified_folder):
      modified_list = list(filename2)
  compared = list(set(original_list) & set (modified_list))
  error_list = []
  finished_list = 0
  for i in compared:
    name_file = os.path.join(original_folder, (compared[compared.index(i)][:-5]))
    original_file = os.path.join(original_folder, (compared[compared.index(i)]))
    modified_file = os.path.join(modified_folder, (compared[compared.index(i)]))
    print(compared[compared.index(i)])
    # some file checks
    if not path.exists(original_file):
      die('Original file does not exist')
    if not path.exists(modified_file):
      die('Modified file does not exist')
    cmp_file = name_file +'_cmp_'+'.docx' # use input filenames, but strip extension
    if path.exists('' + cmp_file + ''):
      print('Comparison file already exists... aborting\nRemove or rename '+cmp_file)
      continue
    # actual Word automation
    #app = wordapp.EnsureDispatch("Word.Application")
    try:
      wordapp.CompareDocuments(wordapp.Documents.Open(original_file), wordapp.Documents.Open(modified_file))
      wordapp.ActiveDocument.ActiveWindow.View.Type = 3  # prevent that word opens itself
      wordapp.ActiveDocument.SaveAs('' + cmp_file + '')
      print('Saved comparison as: ' + cmp_file)
      finished_list += 1
    except:
      error_list += [i]
      pass
  wordapp.Quit()
  print('Number of files which were compared: ' + str(finished_list)+ '\nNumber of files which were not compared: '+ str(len(error_list)) +'\nThese files were not compared: ' + str(error_list))

def main():
  if len(argv) != 3:
    die('Usage: wrd_cmp <original_file> <modified_file>')
  cmp(argv[1], argv[2])

if __name__ == '__main__':
  main()

