
def parse_enum(enum_type, item):
    """Try to parse *item* (string or integer) to enum *type*"""
    if isinstance(item, enum_type):
        return item
    try:
        return enum_type[item]
    except KeyError:
        return enum_type(item)


def set_attr_from_dict(obj, attr_name, mapping, key, ignore_errors=False):
    """
    Tries to set attribute *attr_name* of object *obj* with the value of *key* in the dictionary *map*.
    The item *key* is removed from the *map*.
    If no such *key* is in the dict, nothing is done.
    Any TypeErrors or ValueErrors are silently ignored, if *ignore_errors is set.

    Returns true is *key* was present in *map* and the attribute was successfully set. Otherwise False is returned.
    """
    try:
        value = mapping.pop(key)
    except KeyError:
        return False

    try:
        setattr(obj, attr_name, value)
    except (TypeError, ValueError):
        if ignore_errors:
            return False
        raise

    return True


def set_enum_attr_from_dict(obj, attr_name, enum_type, mapping, key, ignore_errors=False):
    """
    Tries to set enum attribute *attr_name* of object *obj* with the value of *key* in the dictionary *map*.
    The value is parsed to the *enum_type* (using `parse_enum` before the attribute is set).
    The item *key* is removed from the *map*.
    If no such *key* is in the dict nothing is done.
    Any TypeErrors or ValueErrors are silently ignored, if *ignore_errors is set.

    Returns true is *key* was present in *map* and the attribute was successfully set. Otherwise False is returned.
    """
    try:
        value = mapping.pop(key)
    except KeyError:
        return False

    try:
        value = parse_enum(enum_type, value)
    except (KeyError, ValueError):
        if ignore_errors:
            return False
        raise

    try:
        setattr(obj, attr_name, value)
    except (TypeError, ValueError):
        if ignore_errors:
            return False
        raise

    return True
