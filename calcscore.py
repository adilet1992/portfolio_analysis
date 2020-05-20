import math

def calcScorePoint(filial, age, srok, kdd, stazh, cred, edu, zp, sem, pol, pros):
    scorecard = ''
    totalcoef = 0
    totalscore = 0
    agescore = 0
    agecoef = 0
    loancoef = 0
    loanscore = 0
    kddcoef = 0
    kddscore = 0
    workcoef = 0
    workscore = 0
    credcoef = 0
    credscore = 0
    knowcoef = 0
    knowscore = 0
    zpcoef = 0
    zpscore = 0
    polcoef = 0
    polscore = 0
    proscoef = 0
    prosscore = 0
    decision = ''
    pd = 0

    if filial == 'Алматы' or filial == 'Шымкент' or filial == 'Павлодар' or filial == 'Каскелен' or filial == 'Тараз':
        scorecard = 'Юг'
    if filial == 'Нур-Султан' or filial == 'Караганда':
        scorecard = 'Регион'
    if filial == 'Актау' or filial =='Кокшетау':
        scorecard = 'Лоял'

    #################СКОРКАРТА ЮГ#########################################    

    if scorecard == 'Юг':
        if age <= 23:
            agescore = -13
            agecoef = -0.45
        if age > 23 and age <= 27:
            agescore = -6
            agecoef = -0.22
        if age > 27 and age <= 33:
            agescore = 0
            agecoef = 0
        if age > 33 and age <= 45:
            agescore = 6
            agecoef = 0.21
        if age > 45:
            agescore = 13
            agecoef = 0.44

        if srok <= 24:
            loancoef = 0.95
            loanscore = 27
        if srok > 24 and srok <= 36:
            loancoef = 0.49
            loanscore = 14
        if srok > 36 and srok <= 48:
            loancoef = 0
            loanscore = 0
        if srok > 48:
            loancoef = -0.7
            loanscore = -20

        if kdd <= 0.75:
            kddcoef = 0.18
            kddscore = 5
        if kdd > 0.75 and kdd <= 0.86:
            kddcoef = 0
            kddscore = 0
        if kdd > 0.86 and kdd <= 0.94:
            kddcoef = -0.14
            kddscore = -4
        if kdd > 0.94:
            kddcoef = -0.29
            kddscore = -8

        if stazh <= 30:
            workcoef = -0.4
            workscore = -11
        if stazh > 30 and stazh <= 60:
            workcoef = -0.24
            workscore = -7
        if stazh > 60 and stazh <= 120:
            workcoef = 0
            workscore = 0
        if stazh > 120:
            workcoef = 0.16
            workscore = 5

        if cred <= 1:
            credcoef = -0.18
            credscore = -5
        if cred > 1 and cred <= 4:
            credcoef = 0
            credscore = 0
        if cred > 4:
            credcoef = 0.2
            credscore = 6

        if edu == 'Среднее/не оконченное обр.':
            knowcoef = -0.11
            knowscore = -3
        if edu == 'Средне специальное':
            knowcoef = 0
            knowscore = 0
        if edu == 'Высшее/Ученая степень/MBA/Второе высшее':
            knowcoef = 0.31
            knowscore = 9

        if zp == 'Через кассу предприятия':
            zpcoef = -0.29
            zpscore = -8
        if zp == 'с других БВУ':
            zpcoef = 0
            zpscore = 0
        if zp == 'Участник/сотрудник Tengri Bank':
            zpcoef = 0.73
            zpscore = 21

        if sem == 'Холост/незамужем' and pol == 'Мужской':
            polcoef = -0.39
            polscore = -11
        if (sem == 'Холост/незамужем' and pol == 'Женский') or ((sem == 'Гражданский брак' or sem == 'Разведен/разведена' or sem == 'Вдовец/вдова') and pol == 'Мужской'):
            polcoef = -0.09
            polscore = -3
        if  (sem == 'Женат/замужем' and pol == 'Мужской') or ((sem == 'Гражданский брак' or sem == 'Разведен/разведена' or sem == 'Вдовец/вдова') and pol == 'Женский'):
            polcoef = 0
            polscore = 0
        if sem == 'Женат/замужем' and pol == 'Женский':
            polcoef = 0.29
            polscore = 8

        if pros <= 0:
            proscoef = 0.59
            prosscore = 17
        if pros >= 1 and pros <= 30:
            proscoef = 0
            prosscore = 0
        if pros >= 31:
            proscoef = -0.56
            prosscore = -16

        totalscore = 567 + agescore + loanscore + kddscore + workscore + credscore + knowscore + zpscore + polscore + prosscore
        totalcoef = 1.05 + agecoef + loancoef + kddcoef + workcoef + credcoef + knowcoef + zpcoef + polcoef + proscoef

        if totalscore < 561:
            decision = 'Отказано'
        if totalscore >= 561 and totalscore <= 569:
            decision = 'Risk 4'
        if totalscore > 569 and totalscore <= 580:
            decision = 'Risk 3'
        if totalscore > 580 and totalscore <= 599:
            decision = 'Risk 2'
        if totalscore > 599:
            decision = 'Risk 1'
        pd = 1/(1 + math.exp(-totalcoef))

    ###########################СКОРКАРТА РЕГИОН######################################

    if scorecard == 'Регион':
        if age <= 24:
            agescore = -13
            agecoef = -0.44
        if age > 24 and age <= 28:
            agescore = -6
            agecoef = -0.2
        if age > 28 and age <= 32:
            agescore = 0
            agecoef = 0
        if age > 32 and age <= 39:
            agescore = 7
            agecoef = 0.25
        if age > 39:
            agescore = 18
            agecoef = 0.64

        if srok <= 18:
            loancoef = 1.04
            loanscore = 30
        if srok > 18 and srok <= 36:
            loancoef = 0.38
            loanscore = 11
        if srok > 36 and srok <= 48:
            loancoef = 0
            loanscore = 0
        if srok > 48:
            loancoef = -0.86
            loanscore = -25

        if kdd <= 0.66:
            kddcoef = 0.16
            kddscore = 5
        if kdd > 0.66 and kdd <= 0.74:
            kddcoef = 0
            kddscore = 0
        if kdd > 0.74:
            kddcoef = -0.28
            kddscore = -8

        if stazh <= 24:
            workcoef = -0.24
            workscore = -7
        if stazh > 24 and stazh <= 60:
            workcoef = 0
            workscore = 0
        if stazh > 60:
            workcoef = 0.3
            workscore = 9

        if edu == 'Среднее/не оконченное обр.':
            knowcoef = -0.18
            knowscore = -5
        if edu == 'Средне специальное':
            knowcoef = 0
            knowscore = 0
        if edu == 'Высшее/Ученая степень/MBA/Второе высшее':
            knowcoef = 0.59
            knowscore = 17

        if zp == 'Через кассу предприятия':
            zpcoef = -0.31
            zpscore = -9
        if zp == 'с других БВУ':
            zpcoef = 0
            zpscore = 0
        if zp == 'Участник/сотрудник Tengri Bank':
            zpcoef = 0.48
            zpscore = 14

        if sem == 'Холост/незамужем' and pol == 'Мужской':
            polcoef = -0.2
            polscore = -6
        if (sem == 'Холост/незамужем' and pol == 'Женский') or ((sem == 'Гражданский брак' or sem == 'Разведен/разведена') and pol == 'Мужской'):
            polcoef = 0
            polscore = 0
        if ((sem == 'Женат/замужем' or sem == 'Вдовец/вдова') and pol == 'Мужской') or ((sem == 'Гражданский брак' or sem == 'Разведен/разведена') and pol == 'Женский'):        
            polcoef = 0.12
            polscore = 3
        if (sem == 'Женат/замужем' or sem == 'Вдовец/вдова') and pol == 'Женский':
            polcoef = 0.28
            polscore = 8

        if pros <= 0:
            proscoef = 0.56
            prosscore = 16
        if pros >= 1 and pros <= 30:
            proscoef = 0
            prosscore = 0
        if pros >= 31:
            proscoef = -0.76
            prosscore = -22

        totalscore = 567 + agescore + loanscore + kddscore + workscore + knowscore + zpscore + polscore + prosscore
        totalcoef = 1.05 + agecoef + loancoef + kddcoef + workcoef + knowcoef + zpcoef + polcoef + proscoef

        if totalscore < 561:
            decision = 'Отказано'
        if totalscore >= 561 and totalscore <= 568:
            decision = 'Risk 4'
        if totalscore > 568 and totalscore <= 581:
            decision = 'Risk 3'
        if totalscore > 581 and totalscore <= 594:
            decision = 'Risk 2'
        if totalscore > 594:
            decision = 'Risk 1'
        pd = 1/(1 + math.exp(-totalcoef))

    ######################################СКОРКАРТА ЛОЯЛ##################################################


    if scorecard == 'Лоял':
        if age <= 24:
            agescore = -19
            agecoef = -0.67
        if age > 24 and age <= 31:
            agescore = -10
            agecoef = -0.36
        if age > 31 and age <= 42:
            agescore = 0
            agecoef = 0
        if age > 42:
            agescore = 11
            agecoef = 0.39

        if srok <= 24:
            loancoef = 0.65
            loanscore = 19
        if srok > 24 and srok <= 48:
            loancoef = 0
            loanscore = 0
        if srok > 48:
            loancoef = -0.93
            loanscore = -27

        if kdd <= 0.56:
            kddcoef = 0.49
            kddscore = 14
        if kdd > 0.56 and kdd <= 0.82:
            kddcoef = 0
            kddscore = 0
        if kdd > 0.82:
            kddcoef = -0.34
            kddscore = -10

        if stazh <= 18:
            workcoef = -0.19
            workscore = -5
        if stazh > 18 and stazh <= 48:
            workcoef = 0
            workscore = 0
        if stazh > 48 and stazh <= 84:
            workcoef = 0.34
            workscore = 10
        if stazh > 84:
            workcoef = 0.51
            workscore = 15

        if edu == 'Среднее/не оконченное обр.':
            knowcoef = -0.13
            knowscore = -4
        if edu == 'Средне специальное':
            knowcoef = 0
            knowscore = 0
        if edu == 'Высшее/Ученая степень/MBA/Второе высшее':
            knowcoef = 0.46
            knowscore = 13

        if zp == 'Через кассу предприятия':
            zpcoef = -0.24
            zpscore = -7
        if zp == 'с других БВУ':
            zpcoef = 0
            zpscore = 0
        if zp == 'Участник/сотрудник Tengri Bank':
            zpcoef = 0.3
            zpscore = 9

        if (sem == 'Холост/незамужем' or sem == 'Гражданский брак' or sem == 'Вдовец/вдова') and pol == 'Мужской':
            polcoef = -0.37
            polscore = -11
        if ((sem == 'Холост/незамужем' or sem == 'Гражданский брак' or sem == 'Вдовец/вдова') and pol == 'Женский') or ((sem == 'Женат/замужем' or sem == 'Разведен/разведена') and pol == 'Мужской'):
            polcoef = 0
            polscore = 0
        if (sem == 'Женат/замужем' or sem == 'Разведен/разведена') and pol == 'Женский':
            polcoef = 0.13
            polscore = 4

        if pros <= 0:
            proscoef = 0.69
            prosscore = 20
        if pros >= 1 and pros <= 30:
            proscoef = 0
            prosscore = 0
        if pros >= 31:
            proscoef = -0.7
            prosscore = -20

        totalscore = 581 + agescore + loanscore + kddscore + workscore + knowscore + zpscore + polscore + prosscore
        totalcoef = 1.53 + agecoef + loancoef + kddcoef + workcoef + knowcoef + zpcoef + polcoef + proscoef

        if totalscore < 561:
            decision = 'Отказано'
        if totalscore >= 561 and totalscore <= 574:
            decision = 'Risk 4'
        if totalscore > 574 and totalscore <= 579:
            decision = 'Risk 3'
        if totalscore > 579 and totalscore <= 596:
            decision = 'Risk 2'
        if totalscore > 596:
            decision = 'Risk 1'
        pd = 1/(1 + math.exp(-totalcoef))
        
    return (totalscore, decision)