from importlib import resources  # Python 3.7+
import sys


from AttachmentsExtractor.Extractor import extract

def main():
    print(len(sys.argv))
    if len(sys.argv) == 3:
        print(extract(sys.argv[1], sys.argv[2]))

    else: 
        print(' missing number of parameters !!')
        


if __name__=="__main__":
    main()
     