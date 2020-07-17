import wordcloud
from wordcloud import WordCloud
import matplotlib.pyplot as plt
txt = open('pro.txt').read()
print (txt)
print(type(txt))
wc= wordcloud.WordCloud(width=400,height=800)
wc =wc.generate(txt)
plt.imshow(wc, interpolation='bilinear')
plt.axis('off')
wc.to_file('a.png')