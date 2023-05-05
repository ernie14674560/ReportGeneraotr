#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import logging
import traceback
import time
from Template_generator import NestedDict
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import Qt, pyqtSignal
import concurrent.futures as cf
from itertools import repeat
from concurrent.futures import ThreadPoolExecutor

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def concurrent_thread_run(function, var_list, func_args, worker_num):
    multi = ConcurrentThreadsGen(function, var_list, worker_num, func_args)
    result = multi.run()
    return result


class ConcurrentProcessesGen:
    def __init__(self, function, var_list, worker_num=3, func_args=tuple(), func_kwargs=None):

        self.result = {}
        self._func = function
        self._var_list = var_list
        self._worker_num = worker_num
        self._func_args = func_args
        if func_kwargs is None:
            self._func_kwargs = {}
        else:
            self._func_kwargs = func_kwargs

    # def _func_args_unpack(self):

    #     self.

    def _unpack_func(self, var, args, kwargs):
        # pass
        self._func(var, *args, **kwargs)
        # pass

    def run(self):
        with cf.ProcessPoolExecutor(max_workers=self._worker_num) as executor:
            # if self._func_kwargs:

            futures = executor.map(self._unpack_func, self._var_list, repeat(self._func_args),
                                   repeat(self._func_kwargs))
            # else:
            #     futures = executor.map(self._unpack_func, self._var_list, repeat(self._func_args))
        self.result = dict(zip(self._var_list, futures))
        while True:
            if len(self.result) == len(self._var_list):
                break
        return self.result


class ConcurrentThreadsGen(ConcurrentProcessesGen):

    def run(self):
        with ThreadPoolExecutor(max_workers=self._worker_num) as executor:
            # if self._func_kwargs:
            futures = executor.map(self._unpack_func, self._var_list, repeat(self._func_args),
                                   repeat(self._func_kwargs))
            # else:
            #     futures = executor.map(self._unpack_func, self._var_list, repeat(self._func_args))
        self.result = dict(zip(self._var_list, futures))
        return self.result


# class ConcurrentThreadsGen:
#     def __init__(self, function, var_list, worker_num=3, func_args=tuple(), func_kwargs=None):
#
#         self.result = {}
#         self._func = function
#         self._var_list = var_list
#         self._worker_num = worker_num
#         self._func_args = func_args
#         if func_kwargs is None:
#             self._func_kwargs = {}
#         else:
#             self._func_kwargs = func_kwargs
#
#     # def _func_args_unpack(self):
#
#     #     self.
#
#     def _unpack_func(self, var, args, kwargs):
#         # pass
#         self._func(var, *args, **kwargs)
#         # pass
#
#     def run(self):
#         futures = []
#         with cf.ThreadPoolExecutor(max_workers=self._worker_num) as executor:
#             # if self._func_kwargs:
#             for v in self._var_list:
#                 futures.append(executor.submit(self._unpack_func, v, self._func_args,
#                                                self._func_kwargs))
#
#             # else:
#             #     futures = executor.map(self._unpack_func, self._var_list, repeat(self._func_args))
#         # wait(futures)
#         self.result = dict(zip(self._var_list, (f.result() for f in cf.as_completed(
#             futures))))
#         # wait(futures)
#         return self.result


# class MultiProcessGenerator:
#     # function, _param_list, group_size, Nworkers, *args
#     # your final result
#     def __init__(self, function, param_list, worker_num=3, func_args=tuple(), func_kwargs=None):
#
#         self.result = {}
#         self._in_queue = mp.Queue()
#         self._out_queue = mp.Queue()
#         self._func = function
#         self._param_list = param_list
#         self._worker_num = worker_num
#         self._func_args = func_args
#         if func_kwargs is None:
#             self._func_kwargs = {}
#         else:
#             self._func_kwargs = func_kwargs
#
#     def _worker(self):
#         # holds when nothing is available, stops when 'STOP' is seen
#         for param in iter(self._in_queue.get, 'STOP'):
#             # do something
#             result = self._func(*param, *self._func_args, **self._func_kwargs)
#             self._out_queue.put({param: result})  # return your result linked to the input
#
#     def run(self):
#         # fill your input
#         for a in self._param_list:
#             self._in_queue.put(a)
#         # stop command at end of input
#         for n in range(self._worker_num):
#             self._in_queue.put('STOP')
#
#         # setup your worker process doing task as specified
#         process = [mp.Process(target=self._worker, daemon=True) for x in
#                    range(self._worker_num)]
#
#         # run processes
#
#         for p in process:
#             p.start()
#
#         # wait for processes to finish
#
#         for p in process:
#             p.join()
#
#         # collect your results from the calculations
#         for a in self._param_list:
#             self.result.update(self._out_queue.get())
#
#         return self.result


# temp = multiprocess_loop_grouped(worker1, _param_list, group_size, Nworkers, *args)
# map = multiprocess_loop_grouped(worker2, _param_list, group_size, Nworkers, *args)


# class AsyncThreadsGenerator:
#     def __init__(self, variables_list, _func, _func_args=tuple(), max_workers=20):
#         """
#         :param variables_list: list contain variables for generating desired concurrent threads
#         :param _func: _func object for generate concurrent threads
#         :param _func_args: constant args input to the _func
#         """
#         self._variables_list = variables_list
#         self._func = _func
#         self._func_args = _func_args
#         self._max_workers = max_workers
#         self._fetching = asyncio.Queue()
#         # self.executor = ProcessPoolExecutor(max_workers=max_workers)
#         # self.loop = asyncio.get_event_loop()
#
#     async def run(self):
#         all_the_coros = asyncio.gather(*[self._worker(i) for i in range(self._max_workers)])
#         for var in self._variables_list:
#             await self._fetching.put(var)
#         for _ in range(self._max_workers):
#             await self._fetching.put(None)
#         await all_the_coros
#
#     async def fetch(self, var):
#         await asyncio.sleep(1)
#         return var
#
#     async def _worker(self, i):
#         while True:
#             var = await self._fetching.get()
#             if var is None:
#                 return
#             _func = await self.fetch(var)
#
#     def process(self, var):
#         return self._func(var, *self._func_args)
#
#
# class AsyncThreadsGenerator:
#     def __init__(self, variables_list, _func, _func_args=tuple(), max_workers=20):
#         """
#         :param variables_list: list contain variables for generating desired concurrent threads
#         :param _func: _func object for generate concurrent threads
#         :param _func_args: constant args input to the _func
#         """
#         self._variables_list = variables_list
#         self._func = _func
#         self._func_args = _func_args
#         self._max_workers = max_workers
#         self._fetching = asyncio.Queue()
#         self.loop = asyncio.get_event_loop()
#
#     async def awaitable_func(self, var):
#         return self._func(var, *self._func_args)
#
#         await asyncio.sleep(5)
#
#     async def _async_thread_sequence(self):
#         """
#         :return: awaitable gather object
#         """
#         sequence = []
#         for var in self._variables_list:
#             sequence.append(self.awaitable_func(var))
#         result = await asyncio.gather(*sequence)
#         return result
#
#     def run(self):
#         self.loop.run_until_complete(self._async_thread_sequence())

def ui_lock_and_release(ui_dict, lock):
    btns = ui_dict.get('buttons', [])
    text_edits = ui_dict.get('text_edits', [])

    for btn in btns:
        btn.setEnabled(not lock)
    for text_edit in text_edits:
        text_edit.setReadOnly(lock)


class ConnectionObj(QtCore.QObject):
    def __init__(self, val=None):
        self.val = val


def async_task(method, args, kwargs, uid, readycb, errorcb=None):
    """
    Asynchronously runs a task

    :param _func method: the method to run in a thread
    :param object uid: a unique identifier for this task (used for verification)
    :param slot readycb: the callback when data is receieved cb(uid, data)
    :param slot errorcb: the callback when there is an error cb(uid, errmsg)

    The uid option is useful when the calling code makes multiple async_task calls
    and the callbacks need some context about what was sent to the async_task method.
    For example, if you use this method to thread a long running database call
    and the user decides they want to cancel it and start a different one, the
    first one may complete before you have a chance to cancel the task.  In that
    case, the "readycb" will be called with the cancelled task's data.  The uid
    can be used to differentiate those two calls (ie. using the sql query).

    :returns: Request instance
    """
    request = Request(method, args, kwargs, uid, readycb, errorcb)
    QtCore.QThreadPool.globalInstance().start(request)
    return request


class Request(QtCore.QRunnable):
    """
    A Qt object that represents an asynchronous task

    :param _func method: the method to call
    :param list args: list of arguments to pass to method
    :param object uid: a unique identifier (used for verification)
    :param slot readycb: the callback used when data is receieved
    :param slot errorcb: the callback used when there is an error

    The uid param is sent to your error and update callbacks as the
    first argument. It's there to verify the data you're returning

    After created it should be used by invoking:

    .. code-block:: python

       task = Request(...)
       QtCore.QThreadPool.globalInstance().start(task)

    """
    INSTANCES = []
    FINISHED = []

    def __init__(self, method, args, kwargs, uid, readycb, errorcb=None):
        super(Request, self).__init__()
        self.setAutoDelete(True)
        self.cancelled = False

        self.method = method
        self.args = args
        self.kwargs = kwargs
        self.uid = uid
        self.dataReady = readycb
        self.dataError = errorcb

        Request.INSTANCES.append(self)

        # release all of the finished tasks
        Request.FINISHED = []

    def run(self):
        """
        Method automatically called by Qt when the runnable is ready to run.
        This will run in a separate thread.
        """
        # this allows us to "cancel" queued tasks if needed, should be done
        # on shutdown to prevent the app from hanging
        if self.cancelled:
            self.cleanup()
            return

        # runs in a separate thread, for proper async_task signal/slot behavior
        # the object that emits the signals must be created in this thread.
        # Its not possible to run grabber.moveToThread(QThread.currentThread())
        # so to get this QObject to properly exhibit asynchronous
        # signal and slot behavior it needs to live in the thread that
        # we're running in, creating the object from within this thread
        # is an easy way to do that.
        grabber = Requester()
        grabber.Loaded.connect(self.dataReady, Qt.QueuedConnection)
        if self.dataError is not None:
            grabber.Error.connect(self.dataError, Qt.QueuedConnection)

        try:
            result = self.method(*self.args, **self.kwargs)
            if self.cancelled:
                # cleanup happens in 'finally' statement
                return
            grabber.Loaded.emit(self.uid, result)
        except Exception as error:
            if self.cancelled:
                # cleanup happens in 'finally' statement
                return
            # grabber.Error.emit(self.uid, str(error.args[0]))
            grabber.Error.emit(self.uid, error)
            # print error message
            logger.error(traceback.format_exc())
            print(traceback.format_exc())

        finally:
            # this will run even if one of the above return statements
            # is executed inside of the try/except statement see:
            # https://docs.python.org/2.7/tutorial/errors.html#defining-clean-up-actions
            self.cleanup(grabber)

    def cleanup(self, grabber=None):
        # remove references to any object or method for proper ref counting
        self.method = None
        self.args = None
        self.uid = None
        self.dataReady = None
        self.dataError = None

        if grabber is not None:
            grabber.deleteLater()

        # make sure this python obj gets cleaned up
        self.remove()

    def remove(self):
        try:
            Request.INSTANCES.remove(self)

            # when the next request is created, it will clean this one up
            # this will help us avoid this object being cleaned up
            # when it's still being used
            Request.FINISHED.append(self)
        except ValueError:
            # there might be a race condition on shutdown, when shutdown()
            # is called while the thread is still running and the instance
            # has already been removed from the list
            return

    @staticmethod
    def shutdown():
        for inst in Request.INSTANCES:
            inst.cancelled = True
        Request.INSTANCES = []
        Request.FINISHED = []


class Requester(QtCore.QObject):
    """
    A simple object designed to be used in a separate thread to allow
    for asynchronous data fetching
    """

    #
    # Signals
    #

    Error = QtCore.pyqtSignal(object, object)
    """
    Emitted if the fetch fails for any reason

    :param unicode uid: an id to identify this request
    :param unicode error: the error message
    """

    Loaded = QtCore.pyqtSignal(object, object)
    """
    Emitted whenever data comes back successfully

    :param unicode uid: an id to identify this request
    :param list data: the json list returned from the GET
    """

    NetworkConnectionError = QtCore.pyqtSignal(str)
    """
    Emitted when the task fails due to a network connection error

    :param unicode message: network connection error message
    """

    def __init__(self, parent=None):
        """

        :param _func: async thread _func
        :param args: list of args
        :param parent: None
        """
        super(Requester, self).__init__(parent)


class ThreadObject(QtCore.QObject):
    data_finished = pyqtSignal(object)

    def __init__(self, emit=True, return_call_back=True, parent=None):
        super(ThreadObject, self).__init__(parent)
        self.current_uid = 0
        # self.request = None
        self.return_call_back = return_call_back
        self.emit = emit
        self.tasks_queue = NestedDict()

    work_stopped_signal = pyqtSignal()

    def ready_callback(self, uid, result):
        # if uid != self.current_uid:
        #     return
        if self.return_call_back:
            print("Data ready from process uid %s: return %s" % (uid, result))
        disable_ui = self.tasks_queue[self.current_uid, 'packages']
        if disable_ui is not None:
            ui_lock_and_release(disable_ui, lock=False)
        self.tasks_queue.pop(uid)

        if self.emit:
            # if result is not None:
            self.work_stopped_signal.emit()
            self.data_finished.emit(ConnectionObj(result))  # return (purpose, result) to result_classifier

    def error_callback(self, uid, error):
        # if uid != self.current_uid:
        #     return
        if self.return_call_back:
            print("Data error from %s: %s" % (uid, error))
        if self.emit:
            if error is not None:
                self.data_finished.emit(ConnectionObj(('error msg', error)))
            self.work_stopped_signal.emit()
        disable_ui = self.tasks_queue[self.current_uid, 'packages']
        if disable_ui is not None:
            ui_lock_and_release(disable_ui, lock=False)
        self.tasks_queue.pop(uid)

    def fetch(self, func, args=None, kwargs=None, disable_ui=None):
        """

        :param func: multi thread func
        :param args: args pass to func
        :param kwargs: kwargs pass to func
        :param disable_ui: dict like {'text_edits': [text_edit obj ...], 'buttons': [button obj ...]}
        :return:
        """
        # if self.request is not None:
        #     # cancel any pending requests
        #     self.request.cancelled = True
        #     self.request = None
        if args is None:
            args = []
        if kwargs is None:
            kwargs = {}
        self.current_uid += 1
        self.tasks_queue[self.current_uid, 'packages'] = disable_ui
        if disable_ui is not None:
            ui_lock_and_release(disable_ui, lock=True)
        self.tasks_queue[self.current_uid, 'task'] = async_task(func, args, kwargs, self.current_uid,
                                                                self.ready_callback,
                                                                self.error_callback)


def slow_method(arg1, arg2):
    print(
        "Starting slow method")
    time.sleep(1)
    return arg1 + arg2


# class ThreadObject(QtCore.QObject):
#     data_finished = pyqtSignal(object)
#
#     def __init__(self, emit=True, return_call_back=True, parent=None):
#         super(ThreadObject, self).__init__(parent)
#         self.current_uid = 0
#         self.request = None
#         self.return_call_back = return_call_back
#         self.emit = emit
#
#     work_stopped_signal = pyqtSignal()
#
#     def ready_callback(self, uid, result):
#         if uid != self.current_uid:
#             return
#         if self.return_call_back:
#             print("Data ready from process uid %s: return %s" % (uid, result))
#         if self.emit:
#
#             # if result is not None:
#             self.data_finished.emit(ConnectionObj(result))  # return (purpose, result) to result_classifier
#             self.work_stopped_signal.emit()
#
#     def error_callback(self, uid, error):
#         if uid != self.current_uid:
#             return
#         if self.return_call_back:
#             print("Data error from %s: %s" % (uid, error))
#         if self.emit:
#             if error is not None:
#                 self.data_finished.emit(ConnectionObj(('error msg', error)))
#             self.work_stopped_signal.emit()
#
#     def fetch(self, func, args):
#         if self.request is not None:
#             # cancel any pending requests
#             self.request.cancelled = True
#             self.request = None
#
#         self.current_uid += 1
#         # try:
#         self.request = async_task(func, args, self.current_uid,
#                                   self.ready_callback,
#                                   self.error_callback)

def main():
    import sys

    app = QtWidgets.QApplication(sys.argv)

    obj = ThreadObject()

    dialog = QtWidgets.QDialog()
    layout = QtWidgets.QVBoxLayout(dialog)
    button = QtWidgets.QPushButton("Generate", dialog)
    progress = QtWidgets.QProgressBar(dialog)
    progress.setRange(0, 0)
    layout.addWidget(button)
    layout.addWidget(progress)
    button.clicked.connect(obj.fetch(slow_method, ['args1', 'arg2']))
    dialog.show()

    app.exec_()
    app.deleteLater()  # avoids some QThread messages in the shell on exit
    # cancel all running tasks avoid QThread/QTimer error messages
    # on exit
    Request.shutdown()


# def main1():
#     import time
#     import requests
#     def factorial(name, number):
#         f = 1
#         for i in range(2, number + 1):
#             print(f"Task {name}: Compute factorial({i})...")
#
#             f *= i
#         print(f"Task {name}: factorial({number}) = {f}")
#
#     def send_req(url):
#         t = time.time()
#         print("Send a request at", t - start_time, "seconds.")
#         res = requests.get(url)
#
#         t = time.time()
#         print("Receive a response at", t - start_time, "seconds.")
#
#     web_list = ['https://www.google.com.tw/' for i in range(10)]
#     thread = AsyncThreadsGenerator(web_list, send_req)
#     start_time = time.time()
#     thread.run()
#
#
# def main2():
#     class Crawler:
#
#         def __init__(self, urls, max_workers=2):
#             self.urls = urls
#             # create a queue that only allows a maximum of two items
#             self.fetching = asyncio.Queue()
#             self.max_workers = max_workers
#
#         async def crawl(self):
#             # DON'T await here; start consuming things out of the queue, and
#             # meanwhile execution of this function continues. We'll start two
#             # coroutines for fetching and two coroutines for processing.
#             all_the_coros = asyncio.gather(
#                 *[self._worker(i) for i in range(self.max_workers)])
#
#             # place all URLs on the queue
#             for url in self.urls:
#                 await self.fetching.put(url)
#
#             # now put a bunch of `None`'s in the queue as signals to the _worker_num
#             # that there are no more items in the queue.
#             for _ in range(self.max_workers):
#                 await self.fetching.put(None)
#
#             # now make sure everything is done
#             await all_the_coros
#
#         async def _worker(self, i):
#             while True:
#                 url = await self.fetching.get()
#                 if url is None:
#                     # this coroutine is done; simply return to exit
#                     return
#
#                 print(f'Fetch worker {i} is fetching a URL: {url}')
#                 page = await self.fetch(url)
#                 self.process(page)
#
#         async def fetch(self, url):
#             print("Fetching URL: " + url);
#             await asyncio.sleep(2)
#             return f"the contents of {url}"
#
#         def process(self, page):
#             print("processed page: " + page)
#
#     # main loop
#     c = Crawler(['http://www.google.com', 'http://www.yahoo.com',
#                  'http://www.cnn.com', 'http://www.gamespot.com',
#                  'http://www.facebook.com', 'http://www.evergreen.edu'])
#     asyncio.run(c.crawl())
#
#
# def main3():
#     from time import sleep
#
#     import asyncio
#
#     class Crawler:
#
#         def __init__(self, params, _func=print, _func_args=tuple(), max_workers=20):
#             self._func = _func
#             self._func_args = _func_args
#             self.params = params
#             # create a queue that only allows a maximum of two items
#             self.fetching = asyncio.Queue()
#             self.max_workers = max_workers
#
#         async def crawl(self):
#             # DON'T await here; start consuming things out of the queue, and
#             # meanwhile execution of this function continues. We'll start two
#             # coroutines for fetching and two coroutines for processing.
#             all_the_coroutines = asyncio.gather(
#                 *[self._worker(i) for i in range(self.max_workers)])
#
#             # place all URLs on the queue
#             for params in self.params:
#                 await self.fetching.put(params)
#
#             # now put a bunch of `None`'s in the queue as signals to the _worker_num
#             # that there are no more items in the queue.
#             for _ in range(self.max_workers):
#                 await self.fetching.put(None)
#
#             # now make sure everything is done
#             await all_the_coroutines
#
#         async def _worker(self, i):
#             while True:
#                 params = await self.fetching.get()
#                 if params is None:
#                     # this coroutine is done; simply return to exit
#                     return
#
#                 print(f'Fetch worker {i} is working on params: {params}')
#                 var = await self.fetch(params)
#                 self.process(var)
#
#         async def fetch(self, params):
#             #
#             print("Fetching params: " + params)
#             self._func(params, *self._func_args)
#             await asyncio.sleep(0)
#             return f"the contents of {params}"
#
#         def process(self, var):
#             print("processed var: " + var)
#
#     def factorial(name, number):
#         f = 1
#         for i in range(2, number + 1):
#             print(f"Task {name}: Compute factorial({i})...")
#             f *= i
#         print(f"Task {name}: factorial({number}) = {f}")
#
#     # main loop
#     # c = Crawler(['http://www.google.com', 'http://www.yahoo.com',
#     #              'http://www.cnn.com', 'http://www.gamespot.com',
#     #              'http://www.facebook.com', 'http://www.evergreen.edu'])
#     c = Crawler(['A', 'B', 'C'], factorial, (1000,))
#     asyncio.run(c.crawl())


if __name__ == "__main__":
    main()
