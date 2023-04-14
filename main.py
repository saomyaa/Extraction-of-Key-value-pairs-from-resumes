#### imports 


#### user define module
import commanmodule.wordfilemodule.myword as word
import commanmodule.pdffilemodule.mypdf as pdf

def switch(ch):
    if ch == 1:
        word.callwordprocess()
        print("\n\nText File creation begin ......")
        word.generatewordtextfile()
        #word.get_word_xml(fname)
        #print(word.get_xml_tree(fname))
        word.docx_to_xml()
        # word.docx_to_json()
        word.word_to_json_logic()
        

    elif ch == 2:
        pdf.callpdfprocess()
 
fname = ""

print("1. Process Word File ")
print("2. Process PDF  File ")

ch = int(input("Enter Choice : "))

### call modules -- entry point
switch(ch)




    





    