import functools
import logging


def rgetattr(obj, attrname, *args, **kwargs):
    """ Recursive getattr or callable on a COM object"""
    try:
        log = kwargs.pop("log", True)
        if log:
            logging.debug("<= GET: %s, args=%s, kwargs=%s" %
                         (attrname, args, kwargs))
        result = functools.reduce(getattr, attrname.split('.'), obj)
        return result(*args, **kwargs) if args or kwargs else result
    except:
        logging.error("Attribute error %s" % attrname)
        raise AttributeError("AttributeError: %s" % attrname)


def rsetattr(obj, attrname, value):
    """ https://stackoverflow.com/a/31174427 """
    logging.debug("=> SET: %s = %s" % (attrname, value))
    pre, _, post = attrname.rpartition('.')
    return setattr(rgetattr(obj, pre, log=False) if pre else obj, post, value)
