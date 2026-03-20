import UnicornScene from 'unicornstudio-react'

export function SceneCanvas() {
  return (
    <div className="scene-shell" aria-hidden="true">
      <UnicornScene
        projectId="OsK8MprrqJMi8SvUXsD6"
        width="100%"
        height="100%"
        scale={1}
        dpi={1.35}
        sdkUrl="/vendor/unicornStudio.patched.js"
      />
    </div>
  )
}
