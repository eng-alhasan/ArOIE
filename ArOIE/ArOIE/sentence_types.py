def SenType(V,S,O,A1):
  if V is not None:
    if S is not None and O is None and A1 is None:
        return "VS : V=" + str(V), "S=" + str(S)
    if S is not None and O is not None and A1 is None:
        return "VSO : V=" + str(V), "S=" + str(S), "O=" + str(O)
    if S is not None and O is  None and A1 is not None:
        return "VSA : V=" + str(V), "S=" + str(S), "A=" + str(A1)
    if S is not None and O is not None and A1 is not None:
        return "VSOA : V=" + str(V), "S=" + str(S), "O=" + str(O), "A1=" + str(A1)
    if S is None and O is not None and A1 is not None:
        return "VOA : V=" + str(V), "O=" + str(O), "A1=" + str(A1)
    if S is None and O is not None and A1 is None:
        return "VO : V= " + str(V), "O=" + str(O)
    if S is None and O is None and A1 is not None:
        return "VA : V=" + str(V) + "A1=" + str(A1)

def NSenType(I, E):
  if I is not None and E is not None:
      return "IE : I=" + str(I), "E=" + str(E)



