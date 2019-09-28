from django.contrib import admin
from .models import Question,Choice
# Register your models here.

class ChoiceInLine(admin.TabularInline):
    #TabularInline -- display's in tabular way
    #StackedInline
    model = Choice
    extra = 3
class QuestionAdmin(admin.ModelAdmin):
    #fields = ['pub_date','question_text']
    list_display = ['question_text','pub_date','was_published_recently']
    list_filter = ['pub_date']
    search_fields = ['question_text']
    fieldsets = [
    (None,{'fields':['question_text']}),
    ('Date Info',{'fields':['pub_date'],'classes':['collapse']}),
    ]

    inlines = [ChoiceInLine]

admin.site.register(Question,QuestionAdmin)
