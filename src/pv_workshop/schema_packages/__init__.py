from nomad.config.models.plugins import SchemaPackageEntryPoint
from pydantic import Field


class NewSchemaPackageEntryPoint(SchemaPackageEntryPoint):
    parameter: int = Field(0, description='Custom configuration parameter')

    def load(self):
        from pv_workshop.schema_packages.schema_package import m_package

        return m_package


class pg_pvPackageEntryPoint(SchemaPackageEntryPoint):
    parameter: int = Field(0, description='Custom configuration parameter')

    def load(self):
        from pv_workshop.schema_packages.pg_pv_package import m_package

        return m_package


schema_package_entry_point = NewSchemaPackageEntryPoint(
    name='NewSchemaPackage',
    description='New schema package entry point configuration.',
)


pg_pv_schema_package_entry_point = pg_pvPackageEntryPoint(
    name='pg_pvPackage',
    description='pg_pv package entry point configuration.',
)
