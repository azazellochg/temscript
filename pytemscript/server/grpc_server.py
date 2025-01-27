from argparse import Namespace
from concurrent import futures
try:
    import grpc
    import my_grpc_pb2
    import my_grpc_pb2_grpc
except ImportError:
    raise ImportError("Missing dependency 'grpcio', please install it via pip")


class GRPCServer(my_grpc_pb2_grpc.MyServiceServicer):
    def CallMethod(self, request, context):
        method_name = request.method_name
        args = request.args
        kwargs = request.kwargs
        # Call the corresponding method on the server object and return the result
        result = call_method_on_server_object(method_name, *args, **kwargs)
        return my_grpc_pb2.MyResponse(result=result)

def serve(args: Namespace):
    server = grpc.server(futures.ThreadPoolExecutor(max_workers=10))
    my_grpc_pb2_grpc.add_MyServiceServicer_to_server(GRPCServer(), server)
    host = args.host or "::"
    port = args.port or 50051
    useLD = args.useLD
    useTecnaiCCD = args.useTecnaiCCD
    server.add_insecure_port('[%s]:%d' % (host, port))
    server.start()
    server.wait_for_termination()
