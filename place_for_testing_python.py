lst = ['foo', 'bar', 'lorem ipsum', 'color']
start_index = lst.index('foo')
end_index = lst.index('color')
name = lst[start_index+1: end_index]
print(name)