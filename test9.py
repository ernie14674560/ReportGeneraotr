#!/usr/bin/env python
# _*_ coding:utf-8 _*_

def decoratorFunctionWithArguments(arg1, arg2, arg3):
    def wrap(f):
        print("Inside wrap()")

        def wrapped_f(*args):
            print("Inside wrapped_f()")
            print("Decorator arguments:", arg1, arg2, arg3)
            f(*args)
            print("After f(*args)")

        return wrapped_f

    return wrap


@decoratorFunctionWithArguments("hello", "world", 42)
def sayHello(a1, a2, a3, a4):
    print('sayHello arguments:', a1, a2, a3, a4)


def main():
    print("After decoration")
    print("Preparing to call sayHello()")
    sayHello("say", "hello", "argument", "list")
    print("after first sayHello() call")
    sayHello("a", "different", "set of", "arguments")
    print("after second sayHello() call")


if __name__ == '__main__':
    main()
