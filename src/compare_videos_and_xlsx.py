import glob
import xlrd
import re

#pyinstaller compare_videos_and_xlsx.py -n VideosYetToBeProcessed --onefile

program_exit = False
#assumed folders are at same level as program


while program_exit == False:
    
    directory = input("Folder to scan: ")
    print('\n')
    
    #find file with list of videos
    #https://stackoverflow.com/questions/35744613/read-in-xlsx-with-csv-module-in-python
    xlsx_file = glob.glob(directory+"/*.xlsx")
    
    #workaround to avoid temporary file
    for file in xlsx_file:
        if '~$' in file:
            xlsx_file.remove(file)
    
    #if !=1 excel file is found
    while len(xlsx_file)!=1:
        if len(xlsx_file) == 0:
            print("No Excel file found.\n\n")
        if len(xlsx_file) > 1:
            print("Multiple Excel files found, please delete unnecessary Excel files.\n\n")
        directory = input("Folder to scan: ")
        xlsx_file = glob.glob(directory+"/*.xlsx")
        print('\n')
    
    try:
        #initialize excel file
        workbook = xlrd.open_workbook(xlsx_file[0])
        sheet = workbook.sheet_by_index(0)
    except IndexError:
        #in theory this error should never throw
        print("Excel file could not be loaded: IndexError")
        input()
        raise SystemExit(0)
        
    #checks excel formatting
    if sheet.cell_value(0, 22) != 'TITLE':
        print("Improperly formatted Excel file.")
        print('\n\nPress enter to close...')
        input()
        raise SystemExit(0)
    
    #make list of video titles in Excel file
    video_title_from_file = []
    for rowx in range(1,sheet.nrows):
        video_title_from_file.append(sheet.cell_value(rowx, 22))
    
    #find all .mp4 videos
    flist = glob.glob(directory+"/*.mp4")
    
    #found_videos not used but left in in case I need it later
    found_videos=[]
    clean_video_titles = []
    for fname in flist:
        #clean up file name to match excel file
        #some excel files are malformed, this catches those
        try:
            cleanFname = re.search(directory+'\\'+'\(.*).mp4', fname).group(1)
        except AttributeError:
            print("Could not read excel file: AttributeError ln.69.")
            print('\n\nPress enter to close...')
            input()
            raise SystemExit(0)
        #add to clean video list
        clean_video_titles.append(cleanFname)
        #print(cleanFname)
        for title in video_title_from_file:
            #print("comparing "+title+" "+cleanFname)   
            #print(video_title_from_file) 
            if cleanFname == title:
                found_videos.append(title)
                #removes completed videos from list
                clean_video_titles.remove(title)
     
    #print videos that were not in excel file     
    print(directory)
    #make it pretty
    for letter in directory:
        print("-", end="")
    print("")
    for vid in clean_video_titles:
        print(vid)
    #if all videos were found
    if len(clean_video_titles) == 0:
        print('All videos found in Excel file.')
        
    print('\n')
       
    ''' 
    directory = input("\n\nType 'quit' or the name of another folder to scan: ")
    if directory == 'quit':
        program_exit = True
    '''
    
print('\n\nPress enter to close...')
input()