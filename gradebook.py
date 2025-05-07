import tkinter as tk
from tkinter import ttk
import math

class student():
    def __init__(self, name, hw, test):
        self.name = name
        self.hw = hw
        self.tests = test
        self.hlTest = 0
        self.hTest = 0
        self.hwGrade = 0
        self.gradedHw = []
        self.gradedTests = []

    def gradeHw(self):
        self.gradedHw = [self.hw[i] / self.maxHwGrades[i] for i in range(len(self.hw))]
        
    def product(self, a1, a2):
        return [a1[i] * a2[i] for i in range(len(a1))]
    
    def hwAverage(self):
        if 0 in self.hw:
            hwPoints = [self.hw[i] * self.hwWeights[i] for i in range(len(self.hw))]
            maxHwPoints = [self.maxHwGrades[i] * self.hwWeights[i] for i in range(len(self.hw))]
            self.hwGrade = sum(hwPoints) / sum(maxHwPoints)
            return self.hwGrade
        else:
            self.hwGrade = 1
            return 1

    def gradeTests(self):
        self.gradedTests = [self.tests[i] / self.maxTestGrades[i] for i in range(len(self.tests))]
        
    def testAverage(self):
        testScores = [self.tests[i] * self.testWeights[i] for i in range(len(self.tests))]
        maxTestScores = [self.maxTestGrades[i] * self.testWeights[i] for i in range(len(self.maxTestGrades))]
        return sum(testScores) / sum(maxTestScores)

    def hlTestGrade(self):
        tw = self.testWeights
        tw[-1] = 2
        g1 = self.testAverage()
        tw[-1] = 1.5
        g2 = self.testAverage()
        self.hlTest = g1, g2
        self.hTest = max(g1, g2)
        return g1, g2

def arithMean(arr):
    return sum(arr) / len(arr)

def geoMean(arr):
    prod = 1
    for i in arr:
        prod *= i
    return prod ** (1 / len(arr))

def harmMean(arr):
    if 0 in arr:
        return -1
    return len(arr) / sum(map(lambda n: 1 / n, arr))

def median(arr):
    arr.sort()
    n = len(arr) / 2
    a, b = math.floor(n), math.ceil(n)
    return (arr[a] + arr[b]) / 2

def mode(arr):
    s = {}
    for i in arr:
        if i in s:
            s[i] += 1
        else:
            s[i] = 1
    m = 0
    n = -1
    for i in s.items():
        if i[1] > m:
            m = i[1]
            n = i[0]
    return n

def standardDev(arr):
    m = arithMean(arr)
    return (sum([(i - m) ** 2 for i in arr]) / len(arr)) ** 0.5

def dist(students):
    hw = [student.hw for student in students]
    tests = [student.tests for student in students]
    avgHw = [student.hwGrade for student in students]
    avgTests = [student.hTest for student in students]
    
    hw = [[hw[i][j] for i in range(len(hw))] for j in range(len(hw[0]))]
    tests = [[tests[i][j] for i in range(len(tests))] for j in range(len(tests[0]))]
    f = lambda n: (arithMean(n), geoMean(n), harmMean(n), median(n), mode(n), standardDev(n))
    hw = list(map(f, hw))
    tests = list(map(f, tests))
    avgHw = f(avgHw)
    avgTests = f(avgTests)
    return hw, tests, avgHw, avgTests

students = []
with open('data.csv', 'r') as file:
    data = file.read().split(',,,,\n')
    for i in data:
        i = i.split(',,,')
        i = list(map(lambda n: n.split(','), i))
        if not i[0][0]:
            continue
        students.append(student(i[0][0], list(map(int, i[0][1:])), list(map(int, i[1]))))
        students[-1].maxHwGrades = [10, 10, 10, 10, 8, 10, 10, 10, 10, 10]
        students[-1].hwWeights = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
        students[-1].maxTestGrades = [100, 85, 100, 100]
        students[-1].testWeights = [1, 1, 1, 1.5]


for student in students:
    student.gradeHw()
    student.gradeTests()

    print(student.hwAverage())
    print(student.hlTestGrade())
    print('\n')
print(dist(students))
