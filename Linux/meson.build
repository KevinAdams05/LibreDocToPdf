project('doc-to-pdf', 'vala', 'c',
  version: '1.0.0',
  meson_version: '>= 0.62.0',
)

gnome = import('gnome')

dependencies = [
  dependency('glib-2.0'),
  dependency('gobject-2.0'),
  dependency('gio-2.0'),
  dependency('gtk4'),
]

sources = files(
  'src/Application.vala',
  'src/MainWindow.vala',
  'src/AsyncHelpers.vala',
)

executable('doc-to-pdf', sources,
  dependencies: dependencies,
  install: true,
)
