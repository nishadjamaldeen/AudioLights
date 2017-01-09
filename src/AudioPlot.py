from scipy.io import wavfile
from pylab import*
import xlwt
import numpy as np

book= xlwt.Workbook(encoding="utf-8")
sheet = book.add_sheet("Sheet 1")

sampFreq, snd = wavfile.read('Closer_TheChainsmokers.wav')

s1 = snd[:, 0]

timeArray = np.arange(0.000000000000, 23405.0000000000, 23405/10829952.)
timeArray = timeArray/sampFreq
timeArray = timeArray*1000

plot(timeArray, s1, color = 'k')
ylabel('Amplitude')
xlabel('Time (ms)')

a = 0
while (a<80):
    for n, i in enumerate(timeArray):
        if n>a*65535 and n<(a+1)*65535:
            sheet.write(n-a*65535, 3*a, str(i))
    
    for n, i in enumerate(s1):
        if n>a*65535 and n<(a+1)*65535:
            sheet.write(n-a*65535, 3*a+1, str(i))
    a+=1

book.save('Wavefile.xls')