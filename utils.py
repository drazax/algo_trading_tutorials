def chunks(l, n):
    # yield successive n-sized chunks from list l
    for i in range(0, len(l), n):
        yield l[i:i + n]
