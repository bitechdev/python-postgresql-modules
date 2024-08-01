### Copyright (c) 2024 Bitech Systems. All rights reserved.
### The code and materials in this repository are the exclusive property of Bitech Systems and its associated companies and are protected by copyright law.
### Please refer to the license details in the package.

from ..financial.za_vatrate import getVatRate
from ..util.math import btround


def PDAFeeV1(pAmount, pInclusive=True, pVatRate=None):
    if pVatRate is None:
        pVatRate = getVatRate()

    VAT_RATE = pVatRate / 100.0

    if pInclusive:
        # if p_instalment < 13.0:
        #  return 0.0
        if p_instalment <= 200.0:
            return btround(7 * (1 + VAT_RATE))
        elif p_instalment <= 500.0:
            return btround(15 * (1 + VAT_RATE))
        else:
            return btround(25 * (1 + VAT_RATE))

    if not pInclusive:
        if p_instalment < 1.0:
            return 0.0
        elif p_instalment <= btround(200.0 - (7.0 * (1 + VAT_RATE))):
            return btround(7 * (1 + VAT_RATE))
        elif p_instalment <= btround(500.0 - (15.0 * (1 + VAT_RATE))):
            return btround(15 * (1 + VAT_RATE))
        else:
            return btround(25 * (1 + VAT_RATE))

    return 0


def PDAFeeV2(pCollectionAmount, pMaxFee=500, pMinFee=50, pVatRate=None):
    PDAFEE_RATE = 3.0
    if pVatRate is None:
        pVatRate = getVatRate()

    fee = btround(pCollectionAmount * PDAFEE_RATE / 100.0, 2)
    if fee < pMinFee:
        fee = pMinFee
    if fee > pMaxFee:
        fee = pMaxFee

    return fee


def PDAFeeV3(pAmount, pInclusive=True):
    if pInclusive:
        if p_instalment > 500:
            return 15
        elif p_instalment > 200:
            return 10
        elif p_instalment >= 100:
            return 5

    elif not pInclusive:
        if p_instalment > 485:
            return 15
        elif p_instalment > 190:
            return 10
        elif p_instalment >= 95:
            return 5

    return 0


def AftercareFeeCalc(
    pCollectionAmount, pCap=450, pPeriodNr=None, pVatRate=None, pRateStart=5, pRateEnd=3
):

    if pVatRate is None:
        pVatRate = getVatRate()

    rate = pRateStart
    if pPeriodNr is None or pPeriodNr >= 25:
        rate = pRateEnd

    fee = btround(pCollectionAmount * rate / 100.0, 2)
    if fee > pCap:
        fee = pCap

    return fee


def RestrictureFee(pCollectionAmount, pJoined=False):
    Amount = 8000
    AmountJoined = 9000

    if pJoined and pCollectionAmount >= AmountJoined:
        return AmountJoined
    if not pJoined and pCollectionAmount >= Amount:
        return Amount

    return pCollectionAmount


def RegistrationFee(pInclusiveVAT=True, pVatRate=None):
    if pVatRate is None:
        pVatRate = getVatRate()

    if pInclusiveVAT:
        return btround(50 * (pVatRate / 100.0))

    return 50


def AdminFee(pInclusiveVAT=True, pVatRate=None):
    if pVatRate is None:
        pVatRate = getVatRate()

    if pInclusiveVAT:
        return btround(300 * (pVatRate / 100.0))

    return 300


def RecklessFee(pInclusiveVAT=True, pVatRate=None):
    if pVatRate is None:
        pVatRate = getVatRate()

    if pInclusiveVAT:
        return btround(1500 * (pVatRate / 100.0))

    return 1500
