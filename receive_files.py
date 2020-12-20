filename=None
 
if request.method == 'POST' and request.FILES.get('file'):

    from django.core.files.storage import FileSystemStorage

    myfile = request.FILES['file']

    fs = FileSystemStorage()

    filename = fs.save(myfile.name, myfile)
