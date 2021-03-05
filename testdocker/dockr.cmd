S C:\WINDOWS\system32> docker run -it -d -v "c:\temp:/tmp" openxmlpocdocker
1288714e29e4e047cf87fb7e6e7092537ef16410b8c2d8da373f8ae72d7f8a95
PS C:\WINDOWS\system32>


Windows PowerShell
Copyright (C) Microsoft Corporation. Tous droits réservés.

Testez le nouveau système multiplateforme PowerShell https://aka.ms/pscore6

PS C:\WINDOWS\system32> docker

Usage:  docker [OPTIONS] COMMAND

A self-sufficient runtime for containers

Options:
      --config string      Location of client config files (default
                           "C:\\Users\\mohamed.baghdadi\\.docker")
  -c, --context string     Name of the context to use to connect to the
                           daemon (overrides DOCKER_HOST env var and
                           default context set with "docker context use")
  -D, --debug              Enable debug mode
  -H, --host list          Daemon socket(s) to connect to
  -l, --log-level string   Set the logging level
                           ("debug"|"info"|"warn"|"error"|"fatal")
                           (default "info")
      --tls                Use TLS; implied by --tlsverify
      --tlscacert string   Trust certs signed only by this CA (default
                           "C:\\Users\\mohamed.baghdadi\\.docker\\ca.pem")
      --tlscert string     Path to TLS certificate file (default
                           "C:\\Users\\mohamed.baghdadi\\.docker\\cert.pem")
      --tlskey string      Path to TLS key file (default
                           "C:\\Users\\mohamed.baghdadi\\.docker\\key.pem")
      --tlsverify          Use TLS and verify the remote
  -v, --version            Print version information and quit

Management Commands:
  app*        Docker App (Docker Inc., v0.9.1-beta3)
  builder     Manage builds
  buildx*     Build with BuildKit (Docker Inc., v0.5.1-docker)
  config      Manage Docker configs
  container   Manage containers
  context     Manage contexts
  image       Manage images
  manifest    Manage Docker image manifests and manifest lists
  network     Manage networks
  node        Manage Swarm nodes
  plugin      Manage plugins
  scan*       Docker Scan (Docker Inc., v0.5.0)
  secret      Manage Docker secrets
  service     Manage services
  stack       Manage Docker stacks
  swarm       Manage Swarm
  system      Manage Docker
  trust       Manage trust on Docker images
  volume      Manage volumes

Commands:
  attach      Attach local standard input, output, and error streams to a running container
  build       Build an image from a Dockerfile
  commit      Create a new image from a container's changes
  cp          Copy files/folders between a container and the local filesystem
  create      Create a new container
  diff        Inspect changes to files or directories on a container's filesystem
  events      Get real time events from the server
  exec        Run a command in a running container
  export      Export a container's filesystem as a tar archive
  history     Show the history of an image
  images      List images
  import      Import the contents from a tarball to create a filesystem image
  info        Display system-wide information
  inspect     Return low-level information on Docker objects
  kill        Kill one or more running containers
  load        Load an image from a tar archive or STDIN
  login       Log in to a Docker registry
  logout      Log out from a Docker registry
  logs        Fetch the logs of a container
  pause       Pause all processes within one or more containers
  port        List port mappings or a specific mapping for the container
  ps          List containers
  pull        Pull an image or a repository from a registry
  push        Push an image or a repository to a registry
  rename      Rename a container
  restart     Restart one or more containers
  rm          Remove one or more containers
  rmi         Remove one or more images
  run         Run a command in a new container
  save        Save one or more images to a tar archive (streamed to STDOUT by default)
  search      Search the Docker Hub for images
  start       Start one or more stopped containers
  stats       Display a live stream of container(s) resource usage statistics
  stop        Stop one or more running containers
  tag         Create a tag TARGET_IMAGE that refers to SOURCE_IMAGE
  top         Display the running processes of a container
  unpause     Unpause all processes within one or more containers
  update      Update configuration of one or more containers
  version     Show the Docker version information
  wait        Block until one or more containers stop, then print their exit codes

Run 'docker COMMAND --help' for more information on a command.

To get more help with docker, check out our guides at https://docs.docker.com/go/guides/
PS C:\WINDOWS\system32> docker run -it -d -v //c:/temp:/tmp openxmlpocdocker
docker: Error response from daemon: invalid mode: /tmp.
See 'docker run --help'.
PS C:\WINDOWS\system32> docker run -it -d -v "//c:/temp:/tmp" openxmlpocdocker
docker: Error response from daemon: invalid mode: /tmp.
See 'docker run --help'.
PS C:\WINDOWS\system32> docker run -it -d -v "c:\temp:/tmp" openxmlpocdocker
1288714e29e4e047cf87fb7e6e7092537ef16410b8c2d8da373f8ae72d7f8a95
PS C:\WINDOWS\system32> docker ps -a
CONTAINER ID   IMAGE                  COMMAND                  CREATED              STATUS                          PORTS     NAMES
1288714e29e4   openxmlpocdocker       "dotnet OpenXmlPocDo…"   About a minute ago   Exited (0) About a minute ago             hardcore_bouman
bcd31a3ee59a   openxmlpocdocker:dev   "tail -f /dev/null"      26 minutes ago       Exited (137) 6 minutes ago                OpenXmlPocDocker
372834a14876   7a5c3d527993           "dotnet OpenXmlPocDo…"   15 hours ago         Exited (0) About an hour ago              openxmlContainer
PS C:\WINDOWS\system32> docker ps -a
CONTAINER ID   IMAGE                  COMMAND                  CREATED          STATUS                          PORTS     NAMES
1288714e29e4   openxmlpocdocker       "dotnet OpenXmlPocDo…"   2 minutes ago    Exited (0) About a minute ago             hardcore_bouman
bcd31a3ee59a   openxmlpocdocker:dev   "tail -f /dev/null"      27 minutes ago   Exited (137) 7 minutes ago                OpenXmlPocDocker
372834a14876   7a5c3d527993           "dotnet OpenXmlPocDo…"   15 hours ago     Exited (0) About an hour ago              openxmlContainer
PS C:\WINDOWS\system32>
PS C:\WINDOWS\system32> docker ps
CONTAINER ID   IMAGE     COMMAND   CREATED   STATUS    PORTS     NAMES
PS C:\WINDOWS\system32> docker exec -it openxmlpocdocker
"docker exec" requires at least 2 arguments.
See 'docker exec --help'.

Usage:  docker exec [OPTIONS] CONTAINER COMMAND [ARG...]

Run a command in a running container
PS C:\WINDOWS\system32> docker exec -it openxmlpocdocker 1288714E29E4
Error: No such container: openxmlpocdocker
PS C:\WINDOWS\system32> docker ps -a
CONTAINER ID   IMAGE                  COMMAND                  CREATED          STATUS                         PORTS     NAMES
1288714e29e4   openxmlpocdocker       "dotnet OpenXmlPocDo…"   10 minutes ago   Exited (0) 10 minutes ago                hardcore_bouman
bcd31a3ee59a   openxmlpocdocker:dev   "tail -f /dev/null"      35 minutes ago   Up About a minute                        OpenXmlPocDocker
372834a14876   7a5c3d527993           "dotnet OpenXmlPocDo…"   15 hours ago     Exited (0) About an hour ago             openxmlContainer
PS C:\WINDOWS\system32> docker exec -it openxmlpocdocker 1288714e29e4
Error: No such container: openxmlpocdocker
PS C:\WINDOWS\system32> docker exec -it 1288714e29e4 bash
Error response from daemon: Container 1288714e29e4e047cf87fb7e6e7092537ef16410b8c2d8da373f8ae72d7f8a95 is not running
PS C:\WINDOWS\system32> docker ps
CONTAINER ID   IMAGE                  COMMAND               CREATED          STATUS         PORTS     NAMES
bcd31a3ee59a   openxmlpocdocker:dev   "tail -f /dev/null"   37 minutes ago   Up 2 minutes             OpenXmlPocDocker
PS C:\WINDOWS\system32> docker exec -it bcd31a3ee59a bash
root@bcd31a3ee59a:/app# ls
Dockerfile  OpenXmlPocDocker.csproj  OpenXmlPocDocker.csproj.user  Program.cs  Properties  UserDetails.cs  bin  obj
root@bcd31a3ee59a:/app# exit
exit
PS C:\WINDOWS\system32> docker exec -it bcd31a3ee59a bash
root@bcd31a3ee59a:/app# ls
Dockerfile  OpenXmlPocDocker.csproj  OpenXmlPocDocker.csproj.user  Program.cs  Properties  UserDetails.cs  bin  obj
root@bcd31a3ee59a:/app#
