from django.shortcuts import get_object_or_404,render
from django.http import HttpResponse,HttpResponseRedirect
from django.urls import reverse
from .models import Question,Choice
from django.views import generic 

# def index(request):
#     latest_questions = Question.objects.order_by('-pub_date')[:5]
#     context = {'latest_questions' : latest_questions}
#     return render(request,'polls/index.html',context)

# def detail(request,question_id):
#     question = get_object_or_404(Question,pk=question_id)
#     return render(request,'polls/detail.html',{'question':question})


def vote(request,question_id):
    question = get_object_or_404(Question,pk = question_id)
    try:
        selected_choice = question.choice_set.get(pk=request.POST['choice'])
    except (KeyError,Choice.DoesNotExist):
        return render(request,'polls/detail.html',{
            'question': question,
            'error_message' : 'Select a choice and submit'
        })
    else:
        selected_choice.votes+=1
        selected_choice.save()

    return HttpResponseRedirect(reverse('polls:results',args=(question_id,)))

class IndexView(generic.ListView):
    template_name = 'polls/index.html'
    context_object_name = 'latest_questions'
    
    def get_queryset(self):
        return Question.objects.order_by('-pub_date')[:5]

class DetailView(generic.DetailView):
    model = Question
    template_name = 'polls/detail.html'

class ResultsView(generic.DetailView):
    model = Question
    template_name = 'polls/results.html'

# def results(request,question_id):
#     question = get_object_or_404(Question,pk=question_id)
#     return render(request, 'polls/results.html', {'question': question})    

 
# Create your views here.
