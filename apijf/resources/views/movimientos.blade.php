@extends('layouts.app')
@section('content')

    <div class="row">
        <div class="col-md-12">
            <table id="table" class="table table-responsive table-sm">


            </table>

        </div>
    </div>
@endsection
@section('scripts')
    <script>
        $(document).ready(function () {
            $('#table').DataTable({
                "serverSide": true,
                "ajax": "{{ url('api/opera') }}",
                bSort: false,
                columns:[
                    {data:'OPERACION',bSearchable: true, bSortable: false }
                ],
                "scrollCollapse": true,
                "info":           true,
                "paging":         true
            });
        });

    </script>
@endsection

























