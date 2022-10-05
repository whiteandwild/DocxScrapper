
import docx as D
import os , re  , sys
import pprint , json , docx2txt

#to do 
def convert(target , nazwa):

    #os.chdir(sys.path[0])
    print('target --->' , target)
    Doc = D.Document(target)

    def createdir(nazwa):
        try:
            os.mkdir(nazwa)
            return True
        except:
            print('Dir already exists')
            return False
    def checkifupper(par):
        for run in par.runs:
            if run.bold:
                return True
        return False


    createdir(nazwa)
    os.chdir(nazwa)

    prevdir = os.path.abspath(os.path.join(os.getcwd(), os.pardir))

    print('Getting images')
    createdir('Images')
    docx2txt.process(target , 'Images')
    rels = {}
    for r in Doc.part.rels.values():
        if isinstance(r._target, D.parts.image.ImagePart):
            rels[r.rId] = os.path.basename(r._target.partname)

    bigdict = {}
    i = 0
    helperdict = {}
    answ = []
    corr = -1



    print('Data analise')
    for paragraph in Doc.paragraphs:
    
        if 'Graphic' in paragraph._p.xml:
            
            # totalnie wiem co tu sie dzieje
            for rId in rels:
                if rId in paragraph._p.xml:
                    
                    helperdict['img'] = str(os.path.join('Images', rels[rId]))
                    
        else:
            if(re.match('^\d+\.' , paragraph.text)):
                    
                    if(answ):
                        helperdict['odp'] = answ
                    
                        bigdict[str(i+1)] = helperdict

                        answ = []

                        i+=1
                    helperdict = {}                                  

                    corr = 0
                    index = paragraph.text.find(".")
                    helperdict['id'] =  paragraph.text[:index+1]
                    helperdict['question'] = paragraph.text[index:]
            else:
                if(paragraph.text != "" and corr!=-1):

                    answ.append(paragraph.text)
                    if(checkifupper(paragraph)): helperdict['correct'] = corr
                    else: corr+=1

    if(helperdict):

        helperdict['odp'] = answ
                    
        bigdict[str(i+1)] = helperdict
    print('Creating files')
    createdir('pytania')
    #pprint.pprint(bigdict)

    i = 0
    bledy =0
    createdir('pytania/bledy')

    for elem in bigdict:
        name = str(i+1) + ".json"
        where = 'pytania'

        if('correct' not in bigdict[elem]):
            bigdict[elem]['correct'] = 'wpisz tu prawidlowa odpowiedz (a=0 , b=1 itd..'
            where += '/bledy'
            bledy += 1
        
        
        with open(os.path.join(where ,name) , 'w' , encoding='utf-8') as f:
            json.dump(bigdict[elem],f , indent=4, sort_keys=True)
            i+=1

    print('Completed')
    return bledy