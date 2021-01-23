def pronoun_link(UPosTag, source, Pre, UPosTagPre):
    if UPosTag == "PRON" and UPosTagPre != "VERB":
        source = Pre + source
        return source
    else:
        return source
