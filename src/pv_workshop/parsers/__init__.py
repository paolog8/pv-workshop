from nomad.config.models.plugins import ParserEntryPoint
from pydantic import Field


class NewParserEntryPoint(ParserEntryPoint):
    parameter: int = Field(0, description='Custom configuration parameter')

    def load(self):
        from pv_workshop.parsers.parser import NewParser

        return NewParser(**self.model_dump())


class pg_pvExperimentParserEntryPoint(ParserEntryPoint):

    def load(self):
        from pv_workshop.parsers.pg_pv_batch_parser import pg_pvExperimentParser

        return pg_pvExperimentParser(**self.model_dump())


class pg_pvParserEntryPoint(ParserEntryPoint):

    def load(self):
        from pv_workshop.parsers.pg_pv_measurement_parser import pg_pvParser

        return pg_pvParser(**self.model_dump())


parser_entry_point = NewParserEntryPoint(
    name='NewParser',
    description='New parser entry point configuration.',
    mainfile_name_re=r'.*\.newmainfilename',
)


pg_pv_experiment_parser_entry_point = pg_pvExperimentParserEntryPoint(
    name='pg_pvExperimentParserEntryPoint',
    description='pg_pv experiment parser entry point configuration.',
    mainfile_name_re='^(.+\.xlsx)$',
    mainfile_mime_re='(application|text|image)/.*',
)


pg_pv_parser_entry_point = pg_pvParserEntryPoint(
    name='pg_pvParserEntryPoint',
    description='pg_pv parser entry point configuration.',
    mainfile_name_re='^.+\.?.+\.((eqe|jv|mppt)\..{1,4})$',
    mainfile_mime_re='(application|text|image)/.*',
)
