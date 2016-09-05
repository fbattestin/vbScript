Class MyClass
      Public Foo
  End Class

Dim X,Y

Do While 1<>0
  Set X = New MyClass
  Set Y = New MyClass
  
  Set X.Foo = Y
  Set Y.Foo = X
  
  Set X = Nothing
  Set Y = Nothing

Loop
