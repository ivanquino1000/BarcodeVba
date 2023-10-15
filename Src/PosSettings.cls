option explicit

private Type TPosition
    x As Double
    y As Double
    offset as Double
    FallBack as String
End Type

private this as TPosition

private sub Class_Initialize()
    this.x = 1
    this.y = 1
    this.offset = 0
    this.FallBack = "Default"
end sub

public property get X() as Double
    X = this.x
end property

public property let X(byval value as Double)
    this.x = value
end property

public property get Y() as Double
    Y = this.y
end property

public property let Y(byval value as Double)
    this.y = value
end property

public property get Offset() as Double
    Offset = this.offset
end property

public property let Offset(byval value as Double)
    this.offset = value
end property

public property get FallBack() as String
    FallBack = this.FallBack
end property

public property let FallBack(byval value as String)
    this.FallBack = value
end property


