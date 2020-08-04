from importlib import resources  # Python 3.7+
import sys


from pythonApp.oleParser import read_zipped_xml_bin_embeddings

def main():
    print(len(sys.argv))
    if len(sys.argv) == 3:
        print(read_zipped_xml_bin_embeddings(sys.argv[1], sys.argv[2]))

    else: 
        print(' missing number of parameters !!')
        


if __name__=="__main__":
    main()
     