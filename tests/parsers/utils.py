import os

from nomad.client import parse


def set_monkey_patch(monkeypatch):
    def mockreturn_search(*args, upload_id=None):
        return None

    monkeypatch.setattr(
        'pv_workshop.parsers.pg_pv_measurement_parser.set_sample_reference',
        mockreturn_search,
    )

    monkeypatch.setattr(
        'pv_workshop.schema_packages.pg_pv_package.set_sample_reference',
        mockreturn_search,
    )


def delete_json():
    for file in os.listdir(os.path.join('tests','data')):
        if not file.endswith('archive.json'):
            continue
        os.remove(os.path.join('tests','data', file))


def get_archive(file_base, monkeypatch):
    set_monkey_patch(monkeypatch)
    mainfile_path = os.path.join('tests', 'data', file_base)
    file_archive = parse(mainfile_path)[0]
    assert file_archive.data
    assert file_archive.metadata

    for file in os.listdir(os.path.join('tests','data')):
        if 'archive.json' not in file:
            continue
        measurement_path = os.path.join('tests','data', file)
        measurement_archive = parse(measurement_path)[0]

    return measurement_archive
