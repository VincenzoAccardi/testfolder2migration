Public Enum IValassisCouponReturnCode
    OK
    KO
    OK_SKIP_STANDARD
    OK_CONTINUE_STANDARD
    KO_SKIP_STANDARD
    KO_CONTINUE_STANDARD
End Enum

Public Enum IValassisNotificationCancelReasonCode
    ''' <summary>
    '''    Coupon amount > receipt amount
    ''' </summary>
    COUPONGREATHERTHENTOTAMOUNT = 1
    ''' <summary>
    '''    receipt amount lower than requested
    ''' </summary>
    TOTAMOUNTLOWERTHENREQUESTED = 2
    ''' <summary>
    '''    loyalty card not read
    ''' </summary>
    CARDNOTREAD = 3
    ''' <summary>
    '''   requested sku not sold
    ''' </summary>
    SKUNOTSOLD = 4
    ''' <summary>
    '''    item sale canceled
    ''' </summary>
    SALECANDELED = 5
    ''' <summary>
    '''   receipt aborted
    ''' </summary>
    TRXABORTED = 6
    ''' <summary>
    '''    receipt suspended
    ''' </summary>
    TRXSUSPEND = 7
    ''' <summary>
    '''    used another coupon (better discount)
    ''' </summary>
    UAC_BETTERDISCOUNT = 8
    ''' <summary>
    '''    used another coupon (nearer expiration)
    ''' </summary>
    UAC_NEAREREXP = 9
    ''' <summary>
    '''    coupon type not managed
    ''' </summary>
    COUPONTYPENOTMANAGED = 10
End Enum

Public Enum IValassisValidationCouponResultCode
    ''' <summary>
    '''    OK
    ''' </summary>
    OK = 0
    ''' <summary>
    '''    coupon expired
    ''' </summary>
    COUPONEXPIRED = 101
    ''' <summary>
    '''    store Or merchant Not enabled
    ''' </summary>
    STOREORMERCHANT_NOTENABLED = 102
    ''' <summary>
    '''  unique coupon locked
    ''' </summary>
    COUPONLOCKED = 104
    ''' <summary>
    '''    unique coupon already used
    ''' </summary>
    COUPONALREADYUSED = 105
    ''' <summary>
    '''   unique id Not found
    ''' </summary>
    IDNOTFOUND = 106
    ''' <summary>
    '''    external coupon Not found
    ''' </summary>
    EXTERNALCOUPONNOTFOUND = 108
    ''' <summary>
    '''    coupon Not found
    ''' </summary>
    COUPONNOTFOUND = 109
End Enum
