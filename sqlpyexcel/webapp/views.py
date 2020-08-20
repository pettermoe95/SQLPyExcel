from django.shortcuts import render, redirect
from sqlpyexcel.pyexcel import all_reports, start_generating

def index(request):
    context = { 'reports': all_reports()}
    return render(request, 'webapp/index.html', context)

def generate(request):
    if request.method == 'POST':
        year = request.POST['year']
        fromPeriod = request.POST['fromPeriod']
        toPeriod = request.POST['toPeriod']
        print(fromPeriod)
        start_generating(year, fromPeriod, toPeriod)